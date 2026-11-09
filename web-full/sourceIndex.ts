// The source paragraph index of a paragraph the editor replaced by a new one: a cut cell selection
// (prosemirror-tables fills the cell with a default paragraph), or any other edit that swaps a paragraph. The new
// paragraph continues the old one, as the paragraph mark does in Word, so it takes the index back, and markers,
// field writes and caret placement still find the field.
import type { Node as PmNode } from '@tiptap/pm/model'
import type { EditorState, Transaction } from '@tiptap/pm/state'
import { Mapping } from '@tiptap/pm/transform'

/** the source indices a document's paragraphs carry, and where the paragraphs without one start */
function sourceIndices(doc: PmNode) {
  const present = new Set<number>()
  const unindexed: number[] = []
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    if (typeof node.attrs.docxPara === 'number') present.add(node.attrs.docxPara)
    else if (node.attrs.docxPara === null && typeof node.attrs.continuesPara !== 'number') unindexed.push(pos)
    return false
  })
  return { present, unindexed }
}

/** a transaction giving each new paragraph the index of the paragraph it replaced, when no other carries
 *  it; null when there is none */
export function restoreSourceIndices(transactions: readonly Transaction[], before: PmNode, state: EditorState) {
  const { present, unindexed } = sourceIndices(state.doc)
  if (unindexed.length === 0) return null
  const back = new Mapping(transactions.flatMap((tr) => tr.mapping.maps)).invert()
  const tr = state.tr
  for (const pos of unindexed) {
    const replaced = before.nodeAt(back.map(pos, -1))
    const index = replaced?.isTextblock ? replaced.attrs.docxPara : null
    if (typeof index !== 'number' || present.has(index)) continue
    present.add(index)
    tr.setNodeAttribute(pos, 'docxPara', index)
  }
  return tr.docChanged ? tr : null
}
