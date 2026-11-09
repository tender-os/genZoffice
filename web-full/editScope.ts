// Review mode can keep the reviewer's edits inside the fields, so a stray keystroke cannot change the form's
// own text. An edit outside the fields is refused whole and reported, and the host offers the way to unlock
// the document. In either scope a block the editor shows whole (a field, a nested table, a text box) is never
// replaced or deleted, only edited in place: a keystroke on the selected block would otherwise swap the
// form's paragraph for the typed text. Content controls take Enter as Word does: a rich-text block control a
// new paragraph (and joins inside it), a multi-line plain-text control or a rich-text one filling its paragraph
// a line break, any other none, not even a line break. Nor is text moved into or out of a control: the save
// would drop or double it, so the keystroke does nothing. Host writes, the app's own refreshes and a document
// opening are not the reviewer's edits, and a document with no field markers has nothing to protect the
// fields from, so it stays editable.
import type { Node as PmNode } from '@tiptap/pm/model'
import {
  Plugin,
  PluginKey,
  TextSelection,
  type EditorState,
  type Transaction,
} from '@tiptap/pm/state'
import type { Decoration, EditorView } from '@tiptap/pm/view'
import type { EditScope, EditorToHost } from './embed-protocol'
import { blockControlKey } from './blockControl'
import { editedRanges, isUserEdit } from './fieldReporting'
import { HOST_WRITE_META } from './fieldWrite'
import { markersKey, replacedWholeDocument } from './markersPlugin'

export const editScopeKey = new PluginKey('voundEditScope')

let scope: EditScope = 'everything'
const BOX_GLYPHS = /[☐☑☒]/g

/** what the reviewer may edit from now on (vound:editScope); it stays across saves and opens */
export const setEditScope = (next: EditScope) => {
  scope = next
  document.documentElement.classList.toggle('vound-fields-only', next === 'fields')
}

export const isLockedToFields = () => scope === 'fields'

const fieldBlocks = (state: EditorState) =>
  markersKey
    .getState(state)
    ?.set.find()
    .filter((decoration) => decoration.spec.isField) ?? []

/** an insertion is inside a block only strictly; a paragraph added between two blocks is not */
const isWithin = (block: Decoration, range: { from: number; to: number }) =>
  range.from === range.to
    ? block.from < range.from && range.from < block.to
    : block.from <= range.from && range.to <= block.to

const wordsOf = (text: string) => text.replace(BOX_GLYPHS, '☐')

/** a box's line may tick the box; its words are the form's own */
const keepsItsWords = (tr: Transaction, block: Decoration) =>
  !block.spec.togglesOnly ||
  wordsOf(tr.before.textBetween(block.from, block.to)) ===
    wordsOf(tr.doc.textBetween(tr.mapping.map(block.from, -1), tr.mapping.map(block.to, 1)))

const staysInFields = (tr: Transaction, blocks: Decoration[]) =>
  editedRanges([tr]).before.every((range) =>
    blocks.some((block) => isWithin(block, range) && keepsItsWords(tr, block)),
  )

/** a block the editor shows whole and edits in place through its own view; a picture may be removed */
const isWholeBlock = (node: PmNode) =>
  node.type.name === 'docProtected' && node.attrs.blockType !== 'image'

/** whether the block at `pos` before the edit is still there after it: changed attributes keep it */
const keeps = (tr: Transaction, node: PmNode, pos: number) => {
  const after = tr.doc.nodeAt(tr.mapping.map(pos, -1))
  return (
    after?.type === node.type &&
    after.attrs.docxIndex === node.attrs.docxIndex &&
    after.attrs.docxPara === node.attrs.docxPara
  )
}

/** whether the edit replaces or deletes a whole block; an insertion beside one leaves it */
const removesWholeBlock = (tr: Transaction) =>
  editedRanges([tr]).before.some((range) => {
    let removed = false
    if (range.from < range.to) {
      tr.before.nodesBetween(range.from, range.to, (node, pos) => {
        removed ||= isWholeBlock(node) && !keeps(tr, node, pos)
        return !removed
      })
    }
    return removed
  })

/** whether the paragraph at `pos` before the edit is still one paragraph over the same stretch after it */
const staysOneParagraph = (tr: Transaction, node: PmNode, pos: number) => {
  const start = tr.mapping.map(pos, 1)
  const after = tr.doc.nodeAt(start)
  return (
    after?.isTextblock === true &&
    start + after.nodeSize === tr.mapping.map(pos + node.nodeSize, -1)
  )
}

/** the rich-text block control a paragraph takes new paragraphs in, as a key, or null: its number in its
 *  cell, its group when it spans several body paragraphs, else its own paragraph's index (a split twin keeps
 *  each) */
const paragraphControlOf = (node: PmNode) => {
  if (node.attrs.controlEnter !== 'paragraph') return null
  return blockControlKey(node) ?? `index ${node.attrs.docxIndex}`
}

/** the rich-text block control of the paragraph around a position after the edit, or null */
const paragraphControlAt = (doc: PmNode, pos: number) => {
  const { parent } = doc.resolve(pos)
  return parent.isTextblock ? paragraphControlOf(parent) : null
}

/** a paragraph whose control keeps it one paragraph: any control but a rich-text block one */
const keepsOneParagraph = (node: PmNode) => ['lineBreak', 'none'].includes(node.attrs.controlEnter)

const LINE_BREAK = 'hardBreak'

const lineBreaksIn = (node: PmNode | null) => {
  let count = 0
  node?.forEach((child) => {
    if (child.type.name === LINE_BREAK) count += 1
  })
  return count
}

/** whether the paragraph at `pos` gained no line break where its control takes none (Shift+Enter, a paste) */
const gainsNoLineBreak = (tr: Transaction, node: PmNode, pos: number) =>
  node.attrs.controlEnter !== 'none' ||
  lineBreaksIn(tr.doc.nodeAt(tr.mapping.map(pos, 1))) <= lineBreaksIn(node)

/** whether the text of the paragraph at `pos` stays in its control after the edit: a paragraph its control
 *  keeps whole stays one paragraph; a rich-text block control's text stays in it and other text stays out */
const keepsItsControl = (tr: Transaction, node: PmNode, pos: number) => {
  if (keepsOneParagraph(node)) {
    return staysOneParagraph(tr, node, pos) && gainsNoLineBreak(tr, node, pos)
  }
  const control = paragraphControlOf(node)
  return (
    paragraphControlAt(tr.doc, tr.mapping.map(pos + 1, -1)) === control &&
    paragraphControlAt(tr.doc, tr.mapping.map(pos + node.nodeSize - 1, 1)) === control
  )
}

/** Enter in a control that takes line breaks, not paragraphs, inserts a line break */
const enterAsLineBreak = (view: EditorView, event: KeyboardEvent) => {
  const plainEnter =
    event.key === 'Enter' && !event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey
  const { $from, $to } = view.state.selection
  const lineBreak = view.state.schema.nodes[LINE_BREAK]
  if (!plainEnter || !lineBreak || $from.parent !== $to.parent) return false
  if ($from.parent.attrs.controlEnter !== 'lineBreak') return false
  view.dispatch(view.state.tr.replaceSelectionWith(lineBreak.create()).scrollIntoView())
  return true
}

/** Select-all while locked to the fields: only the field around the caret, nothing outside one */
function selectFieldOnly(view: EditorView, event: KeyboardEvent) {
  const selectAll =
    event.key.toLowerCase() === 'a' &&
    (event.metaKey || event.ctrlKey) &&
    !event.shiftKey &&
    !event.altKey
  if (!selectAll || !isLockedToFields()) return false
  const blocks = fieldBlocks(view.state)
  if (blocks.length === 0) return false
  const { from, to } = view.state.selection
  const field = blocks.find((block) => block.from <= from && to <= block.to)
  if (field) {
    const { doc } = view.state
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.between(doc.resolve(field.from + 1), doc.resolve(field.to - 1)),
      ),
    )
  }
  return true
}

/** the reviewer's edit that replaces the whole document (select-all and a keystroke, a cut or a paste), told apart
 *  from a document opening (setContent marks it) and the app's own refreshes */
const replacesDocumentByHand = (tr: Transaction) =>
  replacedWholeDocument(tr) &&
  tr.getMeta('preventUpdate') === undefined &&
  tr.getMeta('addToHistory') !== false &&
  !tr.getMeta(HOST_WRITE_META)

/** whether the edit splits or joins a paragraph a control owns, or moves text into or out of a control */
const breaksControl = (tr: Transaction) =>
  editedRanges([tr]).before.some((range) => {
    let broken = false
    tr.before.nodesBetween(
      Math.max(0, range.from - 1),
      Math.min(tr.before.content.size, range.to + 1),
      (node, pos) => {
        broken ||= node.isTextblock && !keepsItsControl(tr, node, pos)
        return !broken && !node.isTextblock
      },
    )
    return broken
  })

/** the plugin refusing edits outside the fields and edits that remove a whole block, reported through `post`; one per review-mode session */
export const editScopePlugin = (post: (msg: EditorToHost) => void) =>
  new Plugin({
    key: editScopeKey,
    props: {
      handleKeyDown: (view, event) => enterAsLineBreak(view, event) || selectFieldOnly(view, event),
    },
    filterTransaction: (tr, state) => {
      if (tr.getMeta('appendedTransaction')) return true
      if (scope === 'fields' && fieldBlocks(state).length > 0 && replacesDocumentByHand(tr)) {
        post({ type: 'vound:editRefused' })
        return false
      }
      if (!isUserEdit(tr)) return true
      if (breaksControl(tr)) return false
      const refused =
        removesWholeBlock(tr) || (scope === 'fields' && !keepsToFields(tr, fieldBlocks(state)))
      if (refused) post({ type: 'vound:editRefused' })
      return !refused
    },
  })

/** a document with no field markers has nothing to protect the fields from */
const keepsToFields = (tr: Transaction, blocks: Decoration[]) =>
  blocks.length === 0 || staysInFields(tr, blocks)
