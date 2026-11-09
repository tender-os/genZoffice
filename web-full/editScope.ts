// Review mode can keep the reviewer's edits inside the fields, so a stray keystroke cannot change the form's
// own text. An edit outside the fields is refused whole and reported, and the host offers the way to unlock
// the document. In either scope a block the editor shows whole (a field, a nested table, a text box) is never
// replaced or deleted, only edited in place: a keystroke on the selected block would otherwise swap the
// form's paragraph for the typed text. Nor is a paragraph holding or wrapped in a content control split or
// joined (Enter, a Backspace or Delete across its edge, a multi-line paste): the save would drop or double the
// control, so the keystroke does nothing. Host writes, the app's own refreshes and a document opening are not the reviewer's
// edits, and a document with no field markers has nothing to protect the fields from, so it stays editable.
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import type { Decoration } from '@tiptap/pm/view'
import type { EditScope, EditorToHost } from './embed-protocol'
import { editedRanges, isUserEdit } from './fieldReporting'
import { markersKey } from './markersPlugin'

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

/** a paragraph holding an inline content control, or one a block-level control wraps */
const holdsControl = (node: PmNode) =>
  node.isTextblock && (node.attrs.holdsControl === true || node.attrs.sdtShell != null)

/** whether the edit splits or joins a paragraph holding a content control */
const splitsControlParagraph = (tr: Transaction) =>
  editedRanges([tr]).before.some((range) => {
    let split = false
    tr.before.nodesBetween(
      Math.max(0, range.from - 1),
      Math.min(tr.before.content.size, range.to + 1),
      (node, pos) => {
        split ||= holdsControl(node) && !staysOneParagraph(tr, node, pos)
        return !split && !node.isTextblock
      },
    )
    return split
  })

/** the plugin refusing edits outside the fields and edits that remove a whole block, reported through `post`; one per review-mode session */
export const editScopePlugin = (post: (msg: EditorToHost) => void) =>
  new Plugin({
    key: editScopeKey,
    filterTransaction: (tr, state) => {
      if (!isUserEdit(tr) || tr.getMeta('appendedTransaction')) return true
      if (splitsControlParagraph(tr)) return false
      const refused =
        removesWholeBlock(tr) || (scope === 'fields' && !keepsToFields(tr, fieldBlocks(state)))
      if (refused) post({ type: 'vound:editRefused' })
      return !refused
    },
  })

/** a document with no field markers has nothing to protect the fields from */
const keepsToFields = (tr: Transaction, blocks: Decoration[]) =>
  blocks.length === 0 || staysInFields(tr, blocks)
