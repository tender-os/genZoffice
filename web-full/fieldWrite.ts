// The host's field writes (vound:setFieldValue) and caret placement (vound:placeCaret), one rule: exact or
// refused. A write replaces the single occurrence of the field's previous text in the visible text of one
// marked block (a paragraph, or the table cell around it), or inserts into an empty cell when there was no
// previous text. Anything else is refused and the document stays untouched. A write is an ordinary editor
// transaction, so it saves the way typing does; it carries HOST_WRITE_META so edit reporting
// (fieldReporting.ts) does not count it as the reviewer's edit.
import type { Node as PmNode } from '@tiptap/pm/model'
import { Selection, TextSelection, type Transaction } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import type { BlockRef, FieldValueRefusal } from './embed-protocol'
import { locateBlock } from './markers'
import { glideIntoView } from './markersPlugin'

/** transaction meta on every host write: not a user edit */
export const HOST_WRITE_META = 'voundHostWrite'
/** the element of a field passthrough's node view that holds its result text (protected-render.ts) */
export const FIELD_TEXT_SELECTOR = '.doc-field-text'

export type FieldValueResult = { ok: true } | { ok: false; reason: FieldValueRefusal }
export type FieldValueWrite = { block: BlockRef; previous: string; value: string }

/** One piece of a block's visible text: a paragraph's text, a field passthrough's result text
 *  (fieldDisplay.left, which the editor saves into the field result), or field text this rule leaves
 *  alone (a table-of-contents line, a chart caption). */
type TextUnit = { kind: 'text' | 'field' | 'fixed'; pos: number; node: PmNode; text: string }

function unitOf(node: PmNode, pos: number): TextUnit {
  if (node.isTextblock) return { kind: 'text', pos, node, text: node.textContent }
  const display = node.attrs.fieldDisplay
  const text = typeof display?.left === 'string' ? display.left : ''
  // only plain text results: the editor also edits TOC lines (title and page number), which this rule
  // refuses to be safe, and never chart captions. A lone field end has no result text for a save to
  // write into (fieldDisplay.fieldEndOnly), so a value written there would be lost
  const editable = display?.kind === 'text' && !node.attrs.chartDisplay && !display.fieldEndOnly
  return { kind: editable ? 'field' : 'fixed', pos, node, text }
}

/** the text units of a block in document order: the block itself, or every paragraph of a cell */
function unitsOf(block: PmNode, pos: number): TextUnit[] {
  if (block.isTextblock || block.isAtom) return [unitOf(block, pos)]
  const units: TextUnit[] = []
  block.descendants((node, offset) => {
    const leaf = node.isTextblock || (node.isBlock && node.isAtom)
    if (leaf) units.push(unitOf(node, pos + 1 + offset))
    return !leaf
  })
  return units
}

/** start offsets of every occurrence, overlapping ones included ("aa" occurs twice in "aaa") */
function occurrences(text: string, needle: string) {
  const starts: number[] = []
  for (let at = text.indexOf(needle); at !== -1; at = text.indexOf(needle, at + 1)) starts.push(at)
  return starts
}

/** the unit holding text offsets from..to of the units joined with line feeds, with the offsets inside it */
function unitAt(units: TextUnit[], from: number, to: number) {
  let start = 0
  for (const unit of units) {
    const end = start + unit.text.length
    if (from >= start && to <= end) return { unit, from: from - start, to: to - start }
    start = end + 1
  }
  return null
}

type TextRun = { offset: number; pos: number; length: number }

/** the text nodes of a paragraph with their text offsets; inline leaves (breaks, images) add no text */
function textRuns(unit: TextUnit) {
  const runs: TextRun[] = []
  let offset = 0
  unit.node.descendants((child, childPos) => {
    if (!child.isText) return true
    runs.push({ offset, pos: unit.pos + 1 + childPos, length: child.text?.length ?? 0 })
    offset += child.text?.length ?? 0
    return false
  })
  return runs
}

/** Document positions of text offsets from..to of a paragraph. The range starts in the run holding the
 *  first character and ends in the run holding the last, so inline leaves at its edges stay. */
function textRange(unit: TextUnit, from: number, to: number) {
  const runs = textRuns(unit)
  const contentStart = unit.pos + 1
  const holdsStart = runs.find((run) => from >= run.offset && from < run.offset + run.length)
  const holdsEnd = runs.find((run) => to > run.offset && to <= run.offset + run.length)
  const last = runs.at(-1)
  const textEnd = last ? last.pos + last.length : contentStart
  const fromPos = holdsStart ? holdsStart.pos + from - holdsStart.offset : textEnd
  const toPos = holdsEnd ? holdsEnd.pos + to - holdsEnd.offset : fromPos
  return { fromPos, toPos }
}

/** A field display with a new result text. The formatted runs describe the old text (the editor's own field
 *  commit drops them too). An empty result is shown as the field's empty line: without it the editor renders
 *  no result element, and the field could not be typed into again. */
function withResult(display: Record<string, unknown>, left: string) {
  const { runs: _staleRuns, emptyLine: _emptyLine, ...kept } = display
  return left === '' ? { ...kept, left, emptyLine: true } : { ...kept, left }
}

function replaceInUnit(tr: Transaction, unit: TextUnit, from: number, to: number, value: string) {
  if (unit.kind === 'field') {
    const display = withResult(unit.node.attrs.fieldDisplay, unit.text.slice(0, from) + value + unit.text.slice(to))
    return tr.setNodeMarkup(unit.pos, undefined, { ...unit.node.attrs, fieldDisplay: display })
  }
  const { fromPos, toPos } = textRange(unit, from, to)
  // insertText('') would deleteRange, which may take the emptied paragraph along with its text
  return value ? tr.insertText(value, fromPos, toPos) : tr.delete(fromPos, toPos)
}

type Planned = { unit: TextUnit; from: number; to: number } | { refused: FieldValueRefusal }

/** The single match of `needle` inside a field passthrough's result. The editor shows a field's result trimmed
 *  (parse-fields.ts), so the whitespace around the host's previous text may be gone; an empty legacy field's
 *  result is whitespace only and shows as an empty result. A match in a field this rule leaves alone (a TOC
 *  line, a lone field end) is found too, so that it is refused as not editable rather than absent. */
function trimmedFieldMatch(units: TextUnit[], needle: string): Planned {
  const fields = units.filter((unit) => unit.kind !== 'text')
  const matches =
    needle === ''
      ? fields.filter((unit) => unit.text === '').map((unit) => ({ unit, from: 0, to: 0 }))
      : fields.flatMap((unit) => occurrences(unit.text, needle).map((from) => ({ unit, from, to: from + needle.length })))
  if (matches.length === 0) return { refused: 'absent' }
  return matches.length === 1 ? matches[0] : { refused: 'ambiguous' }
}

/** where the rule puts the value, or why it refuses */
function plan(units: TextUnit[], previous: string): Planned {
  if (previous === '' && units.every((unit) => unit.text === '')) {
    // an empty block takes the value, a cell in its first paragraph
    return units[0] ? { unit: units[0], from: 0, to: 0 } : { refused: 'notEditable' }
  }
  const joined = units.map((unit) => unit.text).join('\n')
  const starts = previous === '' ? [] : occurrences(joined, previous)
  if (starts.length > 1) return { refused: 'ambiguous' }
  if (starts.length === 1) return unitAt(units, starts[0], starts[0] + previous.length) ?? { refused: 'notEditable' }
  return trimmedFieldMatch(units, previous.trim())
}

/** vound:setFieldValue: write the value into the block, or refuse and leave the document as it is */
export function applyFieldValue(view: EditorView, write: FieldValueWrite): FieldValueResult {
  const block = locateBlock(view.state.doc, write.block)
  if (!block) return { ok: false, reason: 'noBlock' }
  const planned = plan(unitsOf(block.node, block.pos), write.previous)
  if ('refused' in planned) return { ok: false, reason: planned.refused }
  if (planned.unit.kind === 'fixed') return { ok: false, reason: 'notEditable' }
  if (write.previous === write.value) return { ok: true }
  const tr = replaceInUnit(view.state.tr, planned.unit, planned.from, planned.to, write.value)
  view.dispatch(tr.setMeta(HOST_WRITE_META, true))
  return { ok: true }
}

/** a text caret next to a block, never a NodeSelection on it: with one, the next key would replace the block */
const besideBlock = (view: EditorView, unit: TextUnit) =>
  Selection.findFrom(view.state.doc.resolve(unit.pos + unit.node.nodeSize), 1, true) ??
  Selection.findFrom(view.state.doc.resolve(unit.pos), -1, true) ??
  view.state.selection

/** a result that is whitespace only: Word's empty legacy field shows five en spaces, which typing replaces */
const isBlank = (text: string) => text !== '' && text.trim() === ''

/** Into a field passthrough's text the way a double click gets there (extensions.ts,
 *  wireProtectedInteractionMode): the node view starts editing, and the DOM caret goes to the end of the
 *  result, or selects an empty result's placeholder or a blank result so typing replaces it. Typing there
 *  stays in the node view until the editor commits the field; ProseMirror's own selection moves next to
 *  the block. */
function enterFieldEditing(view: EditorView, unit: TextUnit) {
  const dom = view.nodeDOM(unit.pos)
  const text = dom instanceof HTMLElement ? dom.querySelector(FIELD_TEXT_SELECTOR) : null
  if (!(text instanceof HTMLElement)) return false
  text.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true }))
  text.focus()
  const range = text.ownerDocument.createRange()
  range.selectNodeContents(text)
  if (!unit.node.attrs.fieldDisplay.emptyLine && !isBlank(unit.text)) range.collapse(false)
  const selection = text.ownerDocument.getSelection()
  selection?.removeAllRanges()
  selection?.addRange(range)
  view.dispatch(view.state.tr.setSelection(besideBlock(view, unit)))
  return true
}

/** the end of a paragraph's text, or all of it when it is blank and plain text, so typing replaces it */
function caretInParagraph(view: EditorView, unit: TextUnit) {
  const end = unit.pos + unit.node.nodeSize - 1
  let plainText = true
  unit.node.forEach((child) => { plainText &&= child.isText })
  const from = isBlank(unit.text) && plainText ? unit.pos + 1 : end
  return TextSelection.create(view.state.doc, from, end)
}

/** vound:placeCaret: the caret at the end of the block's text (a cell's last paragraph), focused and
 *  scrolled into view, or inside a field passthrough's result; a blank result is selected instead, so
 *  typing replaces it. False when the block is not found or has no text to put a caret in. */
export function placeCaret(view: EditorView, block: BlockRef): boolean {
  const located = locateBlock(view.state.doc, block)
  if (!located) return false
  const last = unitsOf(located.node, located.pos).at(-1)
  if (!last || last.kind === 'fixed') return false
  if (last.kind === 'field') return enterFieldEditing(view, last)
  view.dispatch(view.state.tr.setSelection(caretInParagraph(view, last)))
  const dom = view.nodeDOM(located.pos)
  if (dom instanceof HTMLElement) glideIntoView(dom, 'nearest')
  view.focus()
  return true
}
