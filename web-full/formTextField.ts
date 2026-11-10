// A legacy text form field (FORMTEXT) in the editor is text carrying the formTextField mark, so a reviewer types
// into each field and the save writes each into its own result. Its range lives only in that mark, so an emptied
// field keeps a zero-width filler, and text typed at a field's start or into a blank field takes the field's mark.
import { combineTransactionSteps, getChangedRanges } from '@tiptap/core'
import { EMPTY_CONTROL_TEXT } from '@genoffice/docx-engine'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, TextSelection, type Transaction } from '@tiptap/pm/state'
import { Decoration, type EditorView } from '@tiptap/pm/view'
import { isHostWrite, isUserEdit } from './fieldReporting'

const FIELD_MARK = 'formTextField'
const PLACEHOLDER_MARK = 'ctrlPlaceholder'
const STYLE_MARK = 'docTextStyle'
const EN_SPACE = ' '

const markOf = (node: PmNode, name: string) => node.marks.find((mark) => mark.type.name === name)

/** each legacy text field of the textblock starting at `start`, as a document range */
function fieldsOf(block: PmNode, start: number) {
  const fields: { index: number; from: number; to: number }[] = []
  block.forEach((child, offset) => {
    const mark = markOf(child, FIELD_MARK)
    if (!mark) return
    const from = start + offset
    const last = fields.at(-1)
    if (last?.index === mark.attrs.index && last.to === from) last.to = from + child.nodeSize
    else fields.push({ index: mark.attrs.index, from, to: from + child.nodeSize })
  })
  return fields
}

/** the legacy text field that holds [from, to], its edges included */
export function formTextFieldAt(doc: PmNode, from: number, to: number) {
  const $from = doc.resolve(from)
  if (!$from.parent.isTextblock) return null
  return (
    fieldsOf($from.parent, $from.start()).find((field) => field.from <= from && to <= field.to) ??
    null
  )
}

/** the text of a textblock around its legacy text fields; '' without any field */
export function labelsOf(block: PmNode) {
  if (fieldsOf(block, 0).length === 0) return ''
  let labels = ''
  block.forEach((child) => {
    if (!markOf(child, FIELD_MARK)) labels += child.text ?? child.type.name
  })
  return labels
}

/** a field typing replaces whole: Word's blank of en spaces, default text, or the filler of an emptied field */
function isReplacedWhole(doc: PmNode, field: { from: number; to: number }) {
  let whole = true
  doc.nodesBetween(field.from, field.to, (node) => {
    if (node.isText)
      whole &&= !!markOf(node, PLACEHOLDER_MARK) || node.text === EN_SPACE.repeat(node.nodeSize)
  })
  return whole
}

const typedMarks = (node: PmNode) =>
  node.marks.filter((mark) => mark.type.name !== PLACEHOLDER_MARK)

/** text typed into a blank field or at a field's start stays in that field; elsewhere the editor's own typing keeps its mark */
function typeIntoField(view: EditorView, from: number, to: number, text: string) {
  const { doc, tr, schema } = view.state
  const field = formTextFieldAt(doc, from, to)
  if (!field) return false
  const replaced = isReplacedWhole(doc, field) ? field : { from, to }
  if (replaced.from !== field.from) return false
  const typed = schema.text(text, typedMarks(doc.nodeAt(field.from)!))
  tr.replaceWith(replaced.from, replaced.to, typed)
  view.dispatch(
    tr.setSelection(TextSelection.create(tr.doc, replaced.from + text.length)).scrollIntoView(),
  )
  return true
}

/** the transaction that gives every field the edit emptied a filler to stay a field */
function keepEmptiedFields(
  transactions: readonly Transaction[],
  before: PmNode,
  doc: PmNode,
  tr: Transaction,
) {
  const { mapping } = combineTransactionSteps(before, [...transactions])
  const emptied: { at: number; node: PmNode }[] = []
  for (const { oldRange } of getChangedRanges(combineTransactionSteps(before, [...transactions]))) {
    before.nodesBetween(oldRange.from, oldRange.to, (block, pos) => {
      if (!block.isTextblock) return true
      const after = doc.nodeAt(mapping.map(pos))
      const kept = after?.isTextblock
        ? fieldsOf(after, mapping.map(pos) + 1).map((field) => field.index)
        : []
      for (const field of fieldsOf(block, pos + 1)) {
        if (!kept.includes(field.index))
          emptied.push({ at: mapping.map(field.from), node: before.nodeAt(field.from)! })
      }
      return false
    })
  }
  for (const { at, node } of emptied.sort((a, b) => b.at - a.at)) {
    const placeholder = doc.type.schema.marks[PLACEHOLDER_MARK].create()
    tr.insert(at, doc.type.schema.text(EMPTY_CONTROL_TEXT, placeholder.addToSet(typedMarks(node))))
  }
  return emptied.length > 0 ? tr.setMeta('addToHistory', false) : null
}

/** a hanging line that opens with a field starts at the indent, where the field's highlight begins; display only */
export function hangingFieldDecorations(node: PmNode, pos: number) {
  const first = node.firstChild
  if (Number(node.attrs.indentFirstLine ?? 0) >= 0 || !first || !markOf(first, FIELD_MARK))
    return []
  return [Decoration.node(pos, pos + node.nodeSize, { style: 'text-indent:0' })]
}

export const formTextKey = new PluginKey('voundFormText')

export const formTextPlugin = new Plugin({
  key: formTextKey,
  props: { handleTextInput: typeIntoField },
  appendTransaction: (transactions, oldState, state) =>
    transactions.some(isUserEdit)
      ? keepEmptiedFields(transactions, oldState.doc, state.doc, state.tr)
      : null,
})

/** The transaction that takes the author's fill-here highlight off every field the changes left holding a value.
 *  Appended to the edit, it shares its undo step, so undoing the value brings the highlight back. */
function dropFilledFieldHighlight(
  transactions: readonly Transaction[],
  before: PmNode,
  doc: PmNode,
  tr: Transaction,
) {
  for (const { newRange } of getChangedRanges(combineTransactionSteps(before, [...transactions]))) {
    doc.nodesBetween(newRange.from, newRange.to, (block, pos) => {
      if (!block.isTextblock) return true
      for (const field of fieldsOf(block, pos + 1).filter((each) => !isReplacedWhole(doc, each))) {
        doc.nodesBetween(field.from, field.to, (node, at) => {
          const style = markOf(node, STYLE_MARK)
          if (!node.isText || !style?.attrs.highlight) return
          const unhighlighted = style.type.create({ ...style.attrs, highlight: null })
          tr.addMark(
            Math.max(at, field.from),
            Math.min(at + node.nodeSize, field.to),
            unhighlighted,
          )
        })
      }
      return false
    })
  }
  return tr.docChanged ? tr : null
}

export const filledFieldHighlightKey = new PluginKey('voundFilledFieldHighlight')

export const filledFieldHighlightPlugin = new Plugin({
  key: filledFieldHighlightKey,
  appendTransaction: (transactions, oldState, state) =>
    transactions.some((tr) => isUserEdit(tr) || (tr.docChanged && isHostWrite(tr)))
      ? dropFilledFieldHighlight(transactions, oldState.doc, state.doc, state.tr)
      : null,
})
