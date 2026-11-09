// Text typed over placeholder text (Vound fork) is the field's own text at once, as in Word: it carries neither
// the ctrlPlaceholder mark nor the placeholder's character style. That holds for typing, which replaces the
// placeholder and keeps its inclusive formatting, for a composition (dead keys) that the browser writes into the
// placeholder's span and ProseMirror reads back with its marks, and for a host write. Undo and redo bring back
// what was there.
import { combineTransactionSteps, getChangedRanges } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import { isPlaceholderStyleId, withoutPlaceholderStyle } from '@genoffice/docx-engine'

const PLACEHOLDER = 'ctrlPlaceholder'
const TEXT_STYLE = 'docTextStyle'

/** the run formatting of typed text: the placeholder's without its character style; null when nothing is left */
function typedTextStyle(attrs: Record<string, unknown>) {
  const rawRPr = typeof attrs.rawRPr === 'string' ? withoutPlaceholderStyle(attrs.rawRPr) : null
  const kept = { ...attrs, styleId: null, rawRPr: rawRPr === '<w:rPr></w:rPr>' ? null : rawRPr }
  return Object.values(kept).some((value) => value !== null) ? kept : null
}

/** the text in [from, to) made the field's own: no placeholder mark, no placeholder character style */
function ownText(tr: Transaction, doc: PmNode, from: number, to: number) {
  const { schema } = doc.type
  tr.removeMark(from, to, schema.marks[PLACEHOLDER])
  doc.nodesBetween(from, to, (node, pos) => {
    const style = node.marks.find((mark) => mark.type.name === TEXT_STYLE)
    if (!node.isText || !style || !isPlaceholderStyleId(String(style.attrs.styleId ?? ''))) return
    const start = Math.max(from, pos)
    const end = Math.min(to, pos + node.nodeSize)
    const attrs = typedTextStyle(style.attrs)
    if (attrs) tr.addMark(start, end, schema.marks[TEXT_STYLE].create(attrs))
    else tr.removeMark(start, end, style)
  })
}

/** text an edit put over or into placeholder text, inside one paragraph */
function typedOver(transactions: readonly Transaction[], before: PmNode, state: EditorState) {
  const type = state.schema.marks[PLACEHOLDER]
  return getChangedRanges(combineTransactionSteps(before, [...transactions])).filter(({ oldRange, newRange }) => {
    if (newRange.from === newRange.to) return false
    const $from = state.doc.resolve(newRange.from)
    if (!$from.parent.isTextblock || !$from.sameParent(state.doc.resolve(newRange.to))) return false
    return before.rangeHasMark(oldRange.from, oldRange.to, type) || state.doc.rangeHasMark(newRange.from, newRange.to, type)
  })
}

export const placeholderTypingPlugin = new Plugin({
  key: new PluginKey('placeholderTyping'),
  appendTransaction(transactions, oldState, state) {
    if (!transactions.some((tr) => tr.docChanged) || transactions.some((tr) => tr.getMeta('history$'))) return null
    const ranges = typedOver(transactions, oldState.doc, state)
    if (ranges.length === 0) return null
    const tr = state.tr
    for (const { newRange } of ranges) ownText(tr, state.doc, newRange.from, newRange.to)
    return tr.docChanged ? tr : null
  },
})
