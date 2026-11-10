// Text typed over placeholder text (Vound fork) is the field's own text at once, as in Word: it carries neither
// the ctrlPlaceholder mark nor the placeholder's character style. That holds for typing, which replaces the
// placeholder and keeps its inclusive formatting, for a composition (dead keys) that the browser writes into the
// placeholder's span and ProseMirror reads back with its marks, and for a host write. Undo and redo bring back
// what was there. The zero-width filler an emptied inline control shows (Vound fork, EMPTY_CONTROL_TEXT) goes as
// soon as text comes in over or right beside it, so it can never turn into the field's own text.
import { combineTransactionSteps, getChangedRanges } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import {
  EMPTY_CONTROL_TEXT,
  isPlaceholderStyleId,
  withoutPlaceholderStyle,
} from '@genoffice/docx-engine'

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

const isFiller = (node: PmNode) =>
  node.isText && node.text === EMPTY_CONTROL_TEXT && node.marks.some((mark) => mark.type.name === PLACEHOLDER)

/** whether the paragraph around `from` held an emptied control's filler touching old [from, to] */
function hadFillerAt(before: PmNode, from: number, to: number) {
  const $from = before.resolve(from)
  if (!$from.parent.isTextblock) return false
  const start = $from.start()
  let found = false
  $from.parent.forEach((node, offset) => {
    found ||= isFiller(node) && start + offset <= to && start + offset + node.nodeSize >= from
  })
  return found
}

/** where the filler lies in or right beside new [from, to]: its marked character, else the first bare one */
function fillerPosition(doc: PmNode, from: number, to: number) {
  const $from = doc.resolve(from)
  const first = Math.max($from.start(), from - 1)
  const last = Math.min($from.end(), to + 1)
  const type = doc.type.schema.marks[PLACEHOLDER]
  const candidates: number[] = []
  for (let pos = first; pos < last; pos += 1) {
    if (doc.textBetween(pos, pos + 1) === EMPTY_CONTROL_TEXT) candidates.push(pos)
  }
  return candidates.find((pos) => doc.rangeHasMark(pos, pos + 1, type)) ?? candidates[0]
}

/** the fillers text just came in over or beside, last first so deleting them keeps the earlier positions */
function fillersReplaced(transactions: readonly Transaction[], before: PmNode, state: EditorState) {
  return getChangedRanges(combineTransactionSteps(before, [...transactions]))
    .flatMap(({ oldRange, newRange }) => {
      if (newRange.from === newRange.to || !hadFillerAt(before, oldRange.from, oldRange.to)) return []
      if (!state.doc.resolve(newRange.from).sameParent(state.doc.resolve(newRange.to))) return []
      return fillerPosition(state.doc, newRange.from, newRange.to) ?? []
    })
    .sort((a, b) => b - a)
}

export const placeholderTypingPlugin = new Plugin({
  key: new PluginKey('placeholderTyping'),
  appendTransaction(transactions, oldState, state) {
    if (!transactions.some((tr) => tr.docChanged)) return null
    // the editor taking over a saved file's parse is no typing: a filler or placeholder it brings keeps its mark
    if (transactions.some((tr) => tr.getMeta('addToHistory') === false)) return null
    // an undo or redo brings text back as it was; only a filler beside it goes, which a redo puts back (Vound fork)
    const undone = transactions.some((tr) => tr.getMeta('history$'))
    const ranges = undone ? [] : typedOver(transactions, oldState.doc, state)
    const fillers = fillersReplaced(transactions, oldState.doc, state)
    if (ranges.length === 0 && fillers.length === 0) return null
    const tr = state.tr
    for (const { newRange } of ranges) ownText(tr, state.doc, newRange.from, newRange.to)
    for (const filler of fillers) tr.delete(filler, filler + 1)
    return tr.docChanged ? tr : null
  },
})
