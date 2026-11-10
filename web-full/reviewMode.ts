// Review mode (vound:open with mode 'review'): the host shows its own bar instead of the ribbon, so the
// editor edits like Word without the ribbon: formatting shortcuts and the app's context menu work. Paste
// inserts plain text in the insertion point's formatting; the document-tool shortcuts do nothing
// (reviewShortcuts.ts), and the context menu leaves out formatted Paste, the field commands, New Comment
// and the AI entries (reviewMenu.ts). host.css hides the ribbon, its tab row and the status bar under
// html.vound-review. Edits and caret moves in marked blocks are reported to the host (fieldReporting.ts), and
// hovering a marked block shows its field chip (fieldChip.ts); the selected and hovered fields get a ring
// around all their blocks (fieldRing.ts).
// A form field selected as a whole block (a click on it selects it as an object) cannot be typed, deleted,
// pasted or cut over: that would replace the block, and with it the field, by plain text. Typing into
// Word's empty form field result (en spaces) replaces the blank instead of adding to it, and the blank
// shows no width, so the caret sits at the field's start (the en spaces stay in the document). Placeholder text
// is selected whole when the caret lands in it (placeholderSelect.ts).
import type { Editor } from '@tiptap/core'
import { NodeSelection, Plugin, PluginKey, TextSelection, type EditorState } from '@tiptap/pm/state'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import type { EditorToHost } from './embed-protocol'
import { boxGapDecorations } from './boxGap'
import { boxTickKey, boxTickPlugin } from './boxTick'
import { editScopeKey, editScopePlugin } from './editScope'
import { formTextKey, formTextPlugin, hangingFieldDecorations } from './formTextField'
import { undoCaretKey, undoCaretPlugin } from './undoCaret'
import { mountFieldChip } from './fieldChip'
import { fieldRingKey, fieldRingPlugin } from './fieldRing'
import { fieldReportingKey, fieldReportingPlugin } from './fieldReporting'
import { placeholderSelectKey, placeholderSelectPlugin } from './placeholderSelect'
import { formFieldBlankStart, isEmptyControlFiller } from './placeholderText'
import { clipboardText } from './pasteText'
import { blockedInReview, isMacPlatform, keepsNativeAction } from './reviewShortcuts'

export const reviewModeKey = new PluginKey('voundReviewMode')

/** Plain text at the selection; insertText keeps the stored marks or those at the insertion point. */
function insertPlainText(view: EditorView, text: string) {
  if (text) view.dispatch(view.state.tr.insertText(text).scrollIntoView())
  return true
}

/** a field passthrough (fieldDisplay) selected as a whole block */
export const selectsFieldBlock = (state: EditorState) =>
  state.selection instanceof NodeSelection && state.selection.node.attrs.fieldDisplay != null

/** swallow an event that would replace a selected field block */
function keepFieldBlock(view: EditorView, event: Event) {
  if (!selectsFieldBlock(view.state)) return false
  event.preventDefault()
  return true
}

const deletesSelection = (event: KeyboardEvent) => event.key === 'Backspace' || event.key === 'Delete'

/** Typing into an empty form field result replaces its en spaces, wherever the caret sits at or in them. That is
 *  how Word's empty legacy form field shows in a table cell (plain text holding the field's result, alone or
 *  after a box or a label), but the editor does not mark field text, so any such text matches: en spaces are
 *  not typed by hand. */
function replaceFormFieldBlank(view: EditorView, from: number, to: number, text: string) {
  const $from = view.state.doc.resolve(from)
  if (!$from.sameParent(view.state.doc.resolve(to))) return false
  const start = formFieldBlankStart($from.parent)
  if (start === null || $from.parentOffset < start) return false
  const blank = $from.start() + start
  // the caret goes after the text: mapped from the blank's start, the next key would type in front of it
  const tr = view.state.tr.insertText(text, blank, $from.end())
  view.dispatch(tr.setSelection(TextSelection.create(tr.doc, blank + text.length)))
  return true
}

/** the class host.css gives an emptied content control's stand-in: a box to click */
export const EMPTY_CONTROL_CLASS = 'vound-empty-control'

/** the class host.css gives an empty form field result's en spaces: no width */
export const FORM_FIELD_BLANK_CLASS = 'vound-blank'

// ponytail: walks the whole document on every change; map the old set if long documents type slowly
function blankDecorations(doc: PmNode) {
  const decorations: Decoration[] = []
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    decorations.push(...boxGapDecorations(node, pos), ...hangingFieldDecorations(node, pos))
    const blank = formFieldBlankStart(node)
    if (blank !== null) decorations.push(Decoration.inline(pos + 1 + blank, pos + node.nodeSize - 1, { class: FORM_FIELD_BLANK_CLASS }))
    node.forEach((child, offset) => {
      if (isEmptyControlFiller(child)) decorations.push(Decoration.inline(pos + 1 + offset, pos + 1 + offset + child.nodeSize, { class: EMPTY_CONTROL_CLASS }))
    })
    return false
  })
  return DecorationSet.create(doc, decorations)
}

export const reviewModePlugin = new Plugin({
  key: reviewModeKey,
  state: {
    init: (_config, state) => blankDecorations(state.doc),
    apply: (tr, blanks) => (tr.docChanged ? blankDecorations(tr.doc) : blanks),
  },
  props: {
    decorations: (state) => reviewModeKey.getState(state),
    // The app's own paste handlers are direct editor props, which ProseMirror asks before any plugin,
    // and some of them insert formatted content themselves. The DOM event reaches this plugin first.
    handleDOMEvents: {
      paste: (view, event) => {
        if (keepFieldBlock(view, event)) return true
        if (!view.editable || !event.clipboardData) return false
        event.preventDefault()
        return insertPlainText(view, clipboardText(event.clipboardData))
      },
      cut: keepFieldBlock,
    },
    // A paste that skips the DOM event (view.pasteHTML) meets the app's direct handlePaste first and lands
    // here only when that one declines. Review mode closes the app's own ways in (the paste-options chip is
    // hidden, its context menu has no formatted Paste); this keeps any remaining one plain. Hard breaks become line feeds.
    handlePaste: (view, _event, slice) =>
      selectsFieldBlock(view.state) || insertPlainText(view, slice.content.textBetween(0, slice.content.size, '\n', '\n')),
    handleTextInput: (view, from, to, text) => selectsFieldBlock(view.state) || replaceFormFieldBlank(view, from, to, text),
    handleKeyPress: (view) => selectsFieldBlock(view.state),
    handleKeyDown: (view, event) => deletesSelection(event) && selectsFieldBlock(view.state),
  },
})

// Window capture runs before every other listener: the editors' keymaps (the main editor, header and footer,
// text boxes) and App.tsx's window keydown handler, which acts without looking at defaultPrevented.
function blockShortcut(event: KeyboardEvent) {
  const mac = isMacPlatform()
  if (!blockedInReview(event, mac)) return
  event.stopPropagation()
  if (!keepsNativeAction(event, mac)) event.preventDefault()
}
/** the field chip of the editor in review mode, mounted once per editor view */
let chip: { view: EditorView; unmount: () => void } | null = null

/** ahead of every other plugin: their keymaps and paste handlers would act first otherwise */
const first = (plugin: Plugin, plugins: Plugin[]) => [plugin, ...plugins]

/** Marks the page for host.css, registers the plugins once, blocks the document-tool shortcuts,
 *  reports the reviewer's edits and caret moves in marked blocks through `post`, and mounts the field chip,
 *  whose presses go through `post` too. */
export function enterReviewMode(editor: Editor, post: (msg: EditorToHost) => void): void {
  document.documentElement.classList.add('vound-review')
  // the same listener added twice is one listener
  window.addEventListener('keydown', blockShortcut, true)
  if (!reviewModeKey.get(editor.state)) editor.registerPlugin(reviewModePlugin, first)
  if (!formTextKey.get(editor.state)) editor.registerPlugin(formTextPlugin, first)
  if (!undoCaretKey.get(editor.state)) editor.registerPlugin(undoCaretPlugin)
  if (!fieldReportingKey.get(editor.state)) editor.registerPlugin(fieldReportingPlugin(post), first)
  if (!boxTickKey.get(editor.state)) editor.registerPlugin(boxTickPlugin, first)
  if (!editScopeKey.get(editor.state)) editor.registerPlugin(editScopePlugin(post), first)
  if (!fieldRingKey.get(editor.state)) editor.registerPlugin(fieldRingPlugin)
  if (!placeholderSelectKey.get(editor.state)) editor.registerPlugin(placeholderSelectPlugin)
  if (chip?.view === editor.view) return
  chip?.unmount()
  chip = { view: editor.view, unmount: mountFieldChip(editor.view, post) }
}

/** Back to the plain embedded editor, for a vound:open without review mode. */
export function leaveReviewMode(editor: Editor): void {
  document.documentElement.classList.remove('vound-review')
  window.removeEventListener('keydown', blockShortcut, true)
  if (reviewModeKey.get(editor.state)) editor.unregisterPlugin(reviewModeKey)
  if (formTextKey.get(editor.state)) editor.unregisterPlugin(formTextKey)
  if (undoCaretKey.get(editor.state)) editor.unregisterPlugin(undoCaretKey)
  if (fieldReportingKey.get(editor.state)) editor.unregisterPlugin(fieldReportingKey)
  if (boxTickKey.get(editor.state)) editor.unregisterPlugin(boxTickKey)
  if (editScopeKey.get(editor.state)) editor.unregisterPlugin(editScopeKey)
  if (fieldRingKey.get(editor.state)) editor.unregisterPlugin(fieldRingKey)
  if (placeholderSelectKey.get(editor.state)) editor.unregisterPlugin(placeholderSelectKey)
  chip?.unmount()
  chip = null
}
