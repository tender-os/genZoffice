// A split source paragraph (Vound fork). ProseMirror's split copies a paragraph's attributes, its source index
// (docxPara) included, to the new half, and an index two paragraphs share names neither. After every edit the
// first paragraph keeps the index and one right after it (or after its continuation) becomes its continuation
// (continuesPara), so a field marked on it covers both halves and writes still find one paragraph. Joining the
// halves keeps the first one's attributes, which gives the one paragraph back; a continuation whose first half
// went takes the index over, and a copy of a paragraph elsewhere gets none.
import { Extension } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, type Transaction } from '@tiptap/pm/state'

type Textblock = { pos: number; end: number; node: PmNode }

function textblocksOf(doc: PmNode) {
  const blocks: Textblock[] = []
  doc.descendants((node, pos) => {
    if (node.isTextblock) blocks.push({ pos, end: pos + node.nodeSize, node })
    return !node.isTextblock
  })
  return blocks
}

const indexOf = (node: PmNode) => (typeof node.attrs.docxPara === 'number' ? node.attrs.docxPara : null)
const continuedOf = (node: PmNode) => (typeof node.attrs.continuesPara === 'number' ? node.attrs.continuesPara : null)

/** the attribute changes that keep every source index on one paragraph and its continuations right after it */
function normalize(tr: Transaction, blocks: Textblock[]) {
  const primaries = new Set(blocks.flatMap(({ node }) => indexOf(node) ?? []))
  // where each index's paragraph and its continuations end, so a continuation must start there
  const chainEnd = new Map<number, number>()
  for (const { pos, end, node } of blocks) {
    const index = indexOf(node) ?? continuedOf(node)
    if (index === null) continue
    const follows = chainEnd.get(index) === pos
    if (!chainEnd.has(index) && indexOf(node) === null && primaries.has(index)) {
      tr.setNodeAttribute(pos, 'continuesPara', null)
      continue
    }
    if (!chainEnd.has(index)) {
      if (indexOf(node) === null) tr.setNodeAttribute(pos, 'docxPara', index)
      if (continuedOf(node) !== null) tr.setNodeAttribute(pos, 'continuesPara', null)
      chainEnd.set(index, end)
      continue
    }
    if (indexOf(node) !== null) tr.setNodeAttribute(pos, 'docxPara', null)
    if (follows) {
      if (continuedOf(node) !== index) tr.setNodeAttribute(pos, 'continuesPara', index)
      chainEnd.set(index, end)
    } else if (continuedOf(node) !== null) tr.setNodeAttribute(pos, 'continuesPara', null)
  }
  return tr
}

// ponytail: walks every paragraph on each change; map the touched ranges if long documents type slowly
export const splitSourceParagraphPlugin = new Plugin({
  key: new PluginKey('splitSourceParagraph'),
  appendTransaction(transactions, _oldState, state) {
    if (!transactions.some((tr) => tr.docChanged)) return null
    const tr = normalize(state.tr, textblocksOf(state.doc))
    return tr.docChanged ? tr : null
  },
})

export const SplitSourceParagraph = Extension.create({
  name: 'splitSourceParagraph',
  addProseMirrorPlugins() {
    return [splitSourceParagraphPlugin]
  },
})
