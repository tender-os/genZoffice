// The host's field writes (vound:setFieldValue) and caret placement (vound:placeCaret), one rule: exact or
// refused. A write replaces the single occurrence of the field's previous text in the visible text of one
// marked block (a paragraph, or the table cell around it), or inserts into an empty cell when there was no
// previous text; a gap write replaces what lies between the gap's two anchors, nothing included. Every
// written place is answered with its gap when one is unique. Anything else is refused and the document stays
// untouched. A write is an ordinary editor transaction, so it saves the way typing does; it carries
// HOST_WRITE_META so edit reporting (fieldReporting.ts) does not count it as the reviewer's edit.
import { EMPTY_CONTROL_TEXT } from '@genoffice/docx-engine'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Selection, TextSelection, type Transaction } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import type { BlockRef, FieldGap, FieldValueRefusal } from './embed-protocol'
import { locateBlock } from './markers'
import { glideIntoView } from './markersPlugin'
import { placeholderAt } from './placeholderText'

/** transaction meta on every host write: not a user edit */
export const HOST_WRITE_META = 'voundHostWrite'
/** the element of a field passthrough's node view that holds its result text (protected-render.ts) */
export const FIELD_TEXT_SELECTOR = '.doc-field-text'

export type FieldValueResult = { ok: true; gap?: FieldGap } | { ok: false; reason: FieldValueRefusal }
// replaces: the block holds nothing but this write's form field, so the write takes the whole field, whatever
// the reviewer typed since the host last read `previous`; 'gap' finds the place by its anchors instead
export type FieldValueWrite =
  | { block: BlockRef; previous: string; value: string; replaces?: 'wholeField' }
  | { block: BlockRef; value: string; replaces: 'gap'; gap: FieldGap }

/** the most characters a side of a returned gap takes */
const GAP_CONTEXT_CAP = 32

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

/** a block's visible text, its paragraphs joined by line feeds */
const joinedText = (units: TextUnit[]) => units.map((unit) => unit.text).join('\n')

/** the visible text of a marked block at `pos`, as a write reads it */
export const blockText = (block: PmNode, pos: number) => joinedText(unitsOf(block, pos))

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
  if (previous === '' && units.every((unit) => unit.text === '' || unit.text === EMPTY_CONTROL_TEXT)) {
    // an empty block takes the value, a cell in its first paragraph, over an emptied control's stand-in
    return units[0] ? { unit: units[0], from: 0, to: units[0].text.length } : { refused: 'notEditable' }
  }
  const joined = joinedText(units)
  const starts = previous === '' ? [] : occurrences(joined, previous)
  if (starts.length > 1) return { refused: 'ambiguous' }
  if (starts.length === 1) return unitAt(units, starts[0], starts[0] + previous.length) ?? { refused: 'notEditable' }
  return trimmedFieldMatch(units, previous.trim())
}

/** the whole of a block's one form field: a field node, or the paragraph a field that fills it loads as */
function wholeField(units: TextUnit[]): Planned {
  const fields = units.filter((unit) => unit.kind !== 'text')
  if (fields.length === 1) return { unit: fields[0], from: 0, to: fields[0].text.length }
  if (fields.length > 1 || units.length !== 1) return { refused: 'absent' }
  return { unit: units[0], from: 0, to: units[0].text.length }
}

/** the places a gap names in a block's joined text: after each `before` (or the start), up to the first
 *  `after` from there (or the end) */
function gapMatches(text: string, gap: FieldGap) {
  const starts = gap.before === '' ? [0] : occurrences(text, gap.before).map((at) => at + gap.before.length)
  return starts.flatMap((from) => {
    const to = gap.after === '' ? text.length : text.indexOf(gap.after, from)
    return to === -1 ? [] : [{ from, to }]
  })
}

/** the gap of text offsets from..to: both sides grow a character at a time, until it names that place alone */
function gapAround(text: string, from: number, to: number): FieldGap | undefined {
  for (let size = 0; size <= GAP_CONTEXT_CAP; size++) {
    const gap = { before: text.slice(Math.max(0, from - size), from), after: text.slice(to, to + size) }
    const matches = gapMatches(text, gap)
    if (matches.length === 1 && matches[0].from === from && matches[0].to === to) return gap
  }
  return undefined
}

function planGap(units: TextUnit[], gap: FieldGap): Planned {
  const matches = gapMatches(joinedText(units), gap)
  if (matches.length === 0) return { refused: 'absent' }
  if (matches.length > 1) return { refused: 'ambiguous' }
  return unitAt(units, matches[0].from, matches[0].to) ?? { refused: 'notEditable' }
}

function planText(units: TextUnit[], write: { previous: string; replaces?: 'wholeField' }): Planned {
  const whole = write.replaces === 'wholeField' ? wholeField(units) : null
  return whole && !('refused' in whole) ? whole : plan(units, write.previous)
}

/** the gap of the value once written in place of from..to of the unit */
function writtenGap(units: TextUnit[], unit: TextUnit, from: number, to: number, value: string) {
  const joined = joinedText(units)
  const start = units.slice(0, units.indexOf(unit)).reduce((offset, each) => offset + each.text.length + 1, 0) + from
  const written = joined.slice(0, start) + value + joined.slice(start + to - from)
  return gapAround(written, start, start + value.length)
}

/** vound:setFieldValue: write the value into the block, or refuse and leave the document as it is */
export function applyFieldValue(view: EditorView, write: FieldValueWrite): FieldValueResult {
  const block = locateBlock(view.state.doc, write.block)
  if (!block) return { ok: false, reason: 'noBlock' }
  const units = unitsOf(block.node, block.pos)
  const planned = write.replaces === 'gap' ? planGap(units, write.gap) : planText(units, write)
  if ('refused' in planned) return { ok: false, reason: planned.refused }
  if (planned.unit.kind === 'fixed') return { ok: false, reason: 'notEditable' }
  const gap = writtenGap(units, planned.unit, planned.from, planned.to, write.value)
  const answer: FieldValueResult = gap ? { ok: true, gap } : { ok: true }
  if (planned.unit.text.slice(planned.from, planned.to) === write.value) return answer
  const tr = replaceInUnit(view.state.tr, planned.unit, planned.from, planned.to, write.value)
  view.dispatch(tr.setMeta(HOST_WRITE_META, true))
  return answer
}

/** a text caret next to a block, never a NodeSelection on it: with one, the next key would replace the block */
const besideBlock = (view: EditorView, unit: TextUnit) =>
  Selection.findFrom(view.state.doc.resolve(unit.pos + unit.node.nodeSize), 1, true) ??
  Selection.findFrom(view.state.doc.resolve(unit.pos), -1, true) ??
  view.state.selection

/** a result that is whitespace only: Word's empty legacy field shows five en spaces, which typing replaces */
const isBlank = (text: string) => text !== '' && text.trim() === ''

/** a form field showing its default text (fieldDisplay.defaultText), which typing replaces as a placeholder */
const showsDefaultText = (node: PmNode) => {
  const display = node.attrs.fieldDisplay
  return typeof display?.defaultText === 'string' && display.left === display.defaultText
}

/** Into a field passthrough's text the way a double click gets there (extensions.ts,
 *  wireProtectedInteractionMode): the node view starts editing, and the DOM caret goes to the end of the
 *  result, or selects an empty result's placeholder, a blank result or the field's default text so typing
 *  replaces it. Typing there
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
  if (!unit.node.attrs.fieldDisplay.emptyLine && !isBlank(unit.text) && !showsDefaultText(unit.node)) range.collapse(false)
  const selection = text.ownerDocument.getSelection()
  selection?.removeAllRanges()
  selection?.addRange(range)
  view.dispatch(view.state.tr.setSelection(besideBlock(view, unit)))
  return true
}

/** a click on a field passthrough showing its default text enters it with that text selected; true when it did */
export function enterFieldShowingDefault(view: EditorView, target: EventTarget | null) {
  const text = target instanceof Element ? target.closest(FIELD_TEXT_SELECTOR) : null
  if (!text) return false
  let entered = false
  view.state.doc.descendants((node, pos) => {
    if (entered || !node.isAtom || !showsDefaultText(node) || !view.nodeDOM(pos)?.contains(text)) return !entered
    entered = enterFieldEditing(view, unitOf(node, pos))
    return false
  })
  return entered
}

/** the end of a paragraph's text, or all of it when it is blank and plain text, or the placeholder text it
 *  ends with, so typing replaces it */
function caretInParagraph(view: EditorView, unit: TextUnit) {
  const end = unit.pos + unit.node.nodeSize - 1
  const placeholder = placeholderAt(view.state.doc, end)
  if (placeholder) return TextSelection.create(view.state.doc, placeholder.from, placeholder.to)
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
