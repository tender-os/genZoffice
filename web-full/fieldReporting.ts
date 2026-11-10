// Review mode tells the host what the reviewer does in marked blocks: vound:blockEdited when an edit of
// theirs changes a marked block's text, with each touched field's text right before and right after it (so
// the host sees an undo take the typing back), and vound:caretBlock when the caret enters a marked block or leaves
// one. Tab is the reviewer's "next field": it neither indents nor inserts a tab, and when the caret is in a
// marked block it reports leaving it with Tab. Mod-Enter and Mod-Shift-Backspace accept and reject the field at
// the caret (fieldActionKeys.ts). The host's own writes (fieldWrite.ts, HOST_WRITE_META) and
// the app's bookkeeping transactions are not the reviewer's edits.
//
// A field passthrough (a legacy form field) is typed into inside its node view: those keys and the DOM caret
// never reach ProseMirror (the node view's stopEvent), and the text reaches the document only when the editor
// commits the field. So the caret counts as in the field's block while the DOM selection is inside a field
// being edited, and Tab there is caught before the node view, commits the field and reports leaving.
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import { Mapping, type StepMap } from '@tiptap/pm/transform'
import type { Decoration, EditorView } from '@tiptap/pm/view'
import type { EditorToHost } from './embed-protocol'
import { fieldActionKeys } from './fieldActionKeys'
import { blockText, HOST_WRITE_META } from './fieldWrite'
import { markersKey, replacedWholeDocument } from './markersPlugin'

export const fieldReportingKey = new PluginKey('voundFieldReporting')

/** the class the editor's protected node views carry while their text is being edited (extensions.ts) */
const EDITING_CLASS = 'doc-content-editing'
/** the window event on which the editor commits pending field, table and text box edits (extensions.ts) */
const COMMIT_EDITS_EVENT = 'ai-docs-commit-tables'

type Range = { from: number; to: number }

/** a host write, or a transaction another plugin appended to one */
export const isHostWrite = (tr: Transaction) =>
  !!tr.getMeta(HOST_WRITE_META) || !!tr.getMeta('appendedTransaction')?.getMeta(HOST_WRITE_META)

/** A change the reviewer made. Not: a host write, a document opening (whole-document replace), or a change
 *  kept out of the undo history, which the app uses for its own refreshes (TOC page numbers, pagination
 *  and formatting stamps) and never for typing. */
export const isUserEdit = (tr: Transaction) =>
  tr.docChanged && !isHostWrite(tr) && tr.getMeta('addToHistory') !== false && !replacedWholeDocument(tr)

/** The ranges the reviewer's steps changed, both in the document before all `trs` (where a joined-away
 *  paragraph still carries its marker) and in the document after them. */
export function editedRanges(trs: readonly Transaction[]) {
  const maps: StepMap[] = trs.flatMap((tr) => tr.mapping.maps)
  const before: Range[] = []
  const after: Range[] = []
  let index = 0
  for (const tr of trs) {
    const counted = isUserEdit(tr)
    for (const map of tr.mapping.maps) {
      const earlier = new Mapping(maps.slice(0, index)).invert()
      const later = new Mapping(maps.slice(index + 1))
      index += 1
      if (!counted) continue
      map.forEach((oldStart, oldEnd, newStart, newEnd) => {
        before.push({ from: earlier.map(oldStart, -1), to: earlier.map(oldEnd, 1) })
        after.push({ from: later.map(newStart, -1), to: later.map(newEnd, 1) })
      })
    }
  }
  return { before, after }
}

/** an insertion touches a block only strictly inside it; a paragraph added between two blocks does not */
const touches = (decoration: Decoration, range: Range) =>
  range.from === range.to
    ? decoration.from < range.from && range.from < decoration.to
    : range.from < decoration.to && range.to > decoration.from

const markerDecorationsOf = (state: EditorState) => markersKey.getState(state)?.set.find() ?? []
const unique = (ids: string[]) => [...new Set(ids)]
const slotIdsOf = (decorations: Decoration[]): string[] => unique(decorations.flatMap((decoration) => decoration.spec.slotIds))
const touchedIds = (state: EditorState, ranges: Range[]) =>
  slotIdsOf(markerDecorationsOf(state).filter((decoration) => ranges.some((range) => touches(decoration, range))))

/** the field node view being edited that holds the DOM caret, if any */
function editedField(view: EditorView) {
  const anchor = view.dom.ownerDocument.getSelection()?.anchorNode
  const element = anchor instanceof Element ? anchor : anchor?.parentElement
  const editing = element?.closest(`.${EDITING_CLASS}`)
  return editing && view.dom.contains(editing) ? editing : null
}

/** the slot ids of the marked blocks whose DOM holds this element */
const idsAround = (view: EditorView, element: Element) =>
  slotIdsOf(markerDecorationsOf(view.state).filter((decoration) => view.nodeDOM(decoration.from)?.contains(element)))

/** the slot ids of the marked blocks the caret is in (a marked paragraph inside a marked cell: both) */
export function caretIds(view: EditorView) {
  const field = editedField(view)
  if (field) return idsAround(view, field)
  const { from } = view.state.selection
  return slotIdsOf(markerDecorationsOf(view.state).filter((d) => d.from <= from && from < d.to))
}

/** the visible text of a field's marked blocks in `state`, joined by line feeds; '' once they are gone */
const fieldText = (state: EditorState, id: string) =>
  markerDecorationsOf(state)
    .filter((decoration) => decoration.spec.slotIds.includes(id))
    .map((decoration) => {
      const block = state.doc.nodeAt(decoration.from)
      return block ? blockText(block, decoration.from) : ''
    })
    .join('\n')

const sameIds = (a: string[], b: string[]) => a.length === b.length && a.every((id, i) => id === b[i])

/** a transaction after which the caret may sit in another marked block: the reviewer moved it or typed,
 *  the host placed it, or the markers changed under it. Host writes and app refreshes only map it. */
const mayMoveCaret = (tr: Transaction) =>
  tr.getMeta(markersKey) !== undefined || (tr.selectionSet && !isHostWrite(tr)) || isUserEdit(tr)

const plainTab = (event: KeyboardEvent) => event.key === 'Tab' && !event.metaKey && !event.ctrlKey && !event.altKey

/** the plugin reporting edits and caret moves to the host through `post`; one per review-mode session */
export function fieldReportingPlugin(post: (msg: EditorToHost) => void) {
  // the ids of the block the host last heard the caret is in; [] after it heard the caret left
  let reported: string[] = []
  // a transaction may have moved the caret: checked once the view shows the new state
  let caretMayHaveMoved = false
  // while Tab commits a field, the caret is about to be reported as leaving it
  let leavingField = false
  const actionKeys = fieldActionKeys(post, caretIds)

  const reportEdits = (trs: readonly Transaction[], oldState: EditorState, state: EditorState) => {
    const { before, after } = editedRanges(trs)
    if (after.length === 0) return
    const ids = unique([...touchedIds(state, after), ...touchedIds(oldState, before)])
    if (ids.length === 0) return
    const texts = ids.map((id) => ({ id, before: fieldText(oldState, id), after: fieldText(state, id) }))
    post({ type: 'vound:blockEdited', ids, texts })
  }

  const reportCaret = (view: EditorView) => {
    const ids = caretIds(view)
    if (sameIds(ids, reported)) return
    reported = ids
    post(ids.length > 0 ? { type: 'vound:caretBlock', ids } : { type: 'vound:caretBlock', ids: [], leftWith: 'other' })
  }

  const leaveByTab = (ids: string[]) => {
    if (ids.length === 0) return
    reported = []
    post({ type: 'vound:caretBlock', ids: [], leftWith: 'Tab' })
  }

  /** Tab inside a field being edited: commit the typed text (reported as an edit first), then leave */
  const tabInField = (view: EditorView, event: KeyboardEvent) => {
    const editing = event.target instanceof Element ? event.target.closest(`.${EDITING_CLASS}`) : null
    if (event.key !== 'Tab' || !editing || !view.dom.contains(editing)) return
    // no focus move to the next control, whichever modifier
    event.preventDefault()
    if (!plainTab(event) || event.shiftKey) return
    const ids = idsAround(view, editing)
    leavingField = true
    view.dom.ownerDocument.defaultView?.dispatchEvent(new Event(COMMIT_EDITS_EVENT))
    leavingField = false
    // the commit rebuilt the field's node view; the caret has not moved anywhere the host should hear of
    caretMayHaveMoved = false
    leaveByTab(ids)
  }

  return new Plugin({
    key: fieldReportingKey,
    // used as the hook that sees every transaction with the states around it; it appends nothing
    appendTransaction: (trs, oldState, state) => {
      reportEdits(trs, oldState, state)
      // another document: whatever the host heard about the caret no longer applies
      if (trs.some(replacedWholeDocument)) reported = []
      if (trs.some(mayMoveCaret)) caretMayHaveMoved = true
      return null
    },
    view: (editorView) => {
      // capture on the editor: ahead of the field node view's own keydown handling
      const onKeyDown = (event: KeyboardEvent) => tabInField(editorView, event)
      editorView.dom.addEventListener('keydown', onKeyDown, true)
      return {
        update: (view) => {
          if (!caretMayHaveMoved || leavingField) return
          caretMayHaveMoved = false
          reportCaret(view)
        },
        destroy: () => editorView.dom.removeEventListener('keydown', onKeyDown, true),
      }
    },
    props: {
      handleKeyDown: (view, event) => {
        if (actionKeys(view, event)) return true
        if (!plainTab(event)) return false
        // Shift+Tab is no way back to the previous field: it does nothing
        if (!event.shiftKey) leaveByTab(caretIds(view))
        // no indent, no tab character, no move to the next table cell (which adds a row at the last one)
        return true
      },
    },
  })
}
