// Vound fork: after an embedded save the saved file's parse differs from the live document in bookkeeping
// (source indices, field results), rarely in text. Replacing the whole document for that would lose the undo
// history, caret and scroll, so a save could never run in the background. When the two have the same shape
// the differences are applied node by node, and text the reviewer typed stays the text they can undo.
import { Fragment, type Node as PmNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'

const lineText = (node: PmNode) => node.textBetween(0, node.content.size, '\n', '\n')

function sameShape(live: PmNode, next: PmNode): boolean {
  if (live.type !== next.type) return false
  if (live.isTextblock) return true
  if (live.childCount !== next.childCount) return false
  for (let index = 0; index < live.childCount; index += 1) {
    if (!sameShape(live.child(index), next.child(index))) return false
  }
  return true
}

/** Only the changed part of a paragraph's content, so what surrounds it keeps its undo steps. */
function rewriteInline(tr: Transaction, live: PmNode, next: PmNode, contentPos: number) {
  const start = live.content.findDiffStart(next.content)
  const end = live.content.findDiffEnd(next.content)
  if (start === null || end === null) return
  const overlap = Math.max(0, start - Math.min(end.a, end.b))
  const to = contentPos + end.a + overlap
  const content: Fragment = next.content.cut(start, end.b + overlap)
  tr.replaceWith(tr.mapping.map(contentPos + start), tr.mapping.map(to), content)
}

function rewriteChildren(tr: Transaction, live: PmNode, next: PmNode, contentPos: number) {
  live.forEach((child, offset, index) => rewriteNode(tr, child, next.child(index), contentPos + offset))
}

function rewriteNode(tr: Transaction, live: PmNode, next: PmNode, pos: number) {
  if (live.eq(next)) return
  if (live.isTextblock) {
    if (lineText(live) !== lineText(next)) rewriteInline(tr, live, next, pos + 1)
  } else {
    rewriteChildren(tr, live, next, pos + 1)
  }
  if (!live.hasMarkup(next.type, next.attrs, next.marks)) {
    tr.setNodeMarkup(tr.mapping.map(pos), undefined, next.attrs, next.marks)
  }
}

/** Adds to `tr` the steps from its document to `next`; false when their shapes differ and `tr` is untouched. */
export function rewriteKeepingHistory(tr: Transaction, next: PmNode): boolean {
  if (!sameShape(tr.doc, next)) return false
  rewriteChildren(tr, tr.before, next, 0)
  return true
}
