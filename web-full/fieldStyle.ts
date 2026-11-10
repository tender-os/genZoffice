// The host's Format control (vound:setFieldStyle): the font and size of a field's text. The target names the
// field as a write does (fieldWrite.ts), else as an empty field (fieldStylePlace.ts); the selection is restyled
// when it lies inside the field, else the whole field, and a field passthrough's display (fieldStyleBlock.ts).
// Every restyle is one undoable transaction with HOST_WRITE_META, saved as the run's w:rFonts and w:sz; an empty
// place stores the style for the text typed there next. The answer is the style the text shows now; an empty style
// changes nothing and only answers.
import type { Node as PmNode } from '@tiptap/pm/model'
import { TextSelection } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import type { FieldTextStyle, FieldValueRefusal } from './embed-protocol'
import { fieldBlockElements, restyleFieldBlock } from './fieldStyleBlock'
import { emptyFieldPlace } from './fieldStylePlace'
import { fieldTextPlace, HOST_WRITE_META, type FieldTarget, type FieldTextPlace } from './fieldWrite'

/** `formText`: the index of the legacy text form field in its paragraph, for a field its text does not name once */
export type FieldStyleRequest = FieldTarget & { formText?: number; style: FieldTextStyle }
export type FieldStyleResult = { ok: true; style: FieldTextStyle } | { ok: false; reason: FieldValueRefusal }
type Range = { from: number; to: number }

const PT_PER_PX = 0.75

/** the docTextStyle attributes of a style: the Latin face and the size in half-points */
const markAttrsOf = (style: FieldTextStyle) => ({
  ...(style.fontFamily === undefined ? {} : { fontAscii: style.fontFamily }),
  ...(style.fontSize === undefined ? {} : { sizeHalfPoints: Math.round(style.fontSize * 2) }),
})

/** the selection when it lies inside the field, else the field */
function restyledRange(view: EditorView, field: Range): Range {
  const { from, to, empty } = view.state.selection
  return !empty && from >= field.from && to <= field.to ? { from, to } : field
}

function textSegments(doc: PmNode, { from, to }: Range) {
  const segments: Range[] = []
  doc.nodesBetween(from, to, (node, pos) => {
    if (node.isText) segments.push({ from: Math.max(pos, from), to: Math.min(pos + node.nodeSize, to) })
    return !node.isText
  })
  return segments
}

function restyle(view: EditorView, range: Range, attrs: ReturnType<typeof markAttrsOf>) {
  const { doc, schema, tr } = view.state
  const type = schema.marks.docTextStyle
  if (range.from === range.to) {
    const marks = doc.resolve(range.from).marks()
    const existing = type.isInSet(marks)
    const stored = [...marks.filter((mark) => mark.type !== type), type.create({ ...existing?.attrs, ...attrs })]
    view.dispatch(tr.setSelection(TextSelection.create(doc, range.from)).setStoredMarks(stored))
    view.focus()
    return
  }
  for (const segment of textSegments(doc, range)) {
    const existing = type.isInSet(doc.resolve(segment.from + 1).marks())
    tr.addMark(segment.from, segment.to, type.create({ ...existing?.attrs, ...attrs }))
  }
  view.dispatch(tr.setMeta(HOST_WRITE_META, true))
}

function elementAt(view: EditorView, pos: number) {
  const { node } = view.domAtPos(pos)
  return node instanceof Element ? node : node.parentElement
}

const firstFamily = (css: string) => (css.split(',')[0] ?? '').trim().replaceAll('"', '').replaceAll("'", '')

/** points of a computed font size: px in a browser, the declared pt in jsdom */
function pointsOf(css: string) {
  const value = parseFloat(css)
  if (!Number.isFinite(value)) return undefined
  const points = css.endsWith('px') ? value * PT_PER_PX : value
  return Math.round(points * 2) / 2
}

const only = <T>(values: T[]) => (new Set(values).size === 1 ? values[0] : undefined)

/** the elements that show the range's text, or the one around an empty range */
function textElements(view: EditorView, range: Range) {
  const segments = textSegments(view.state.doc, range)
  const positions = segments.length > 0 ? segments.map((segment) => segment.from + 1) : [range.from]
  return positions.flatMap((pos) => {
    const element = elementAt(view, pos)
    return element ? [element] : []
  })
}

/** the face and size the elements show, each left out when they show more than one */
function shownStyle(elements: Element[]): FieldTextStyle {
  const styles = elements.map((element) => getComputedStyle(element))
  const fontFamily = only(styles.map((style) => firstFamily(style.fontFamily)))
  const fontSize = only(styles.map((style) => pointsOf(style.fontSize)))
  return { ...(fontFamily ? { fontFamily } : {}), ...(fontSize === undefined ? {} : { fontSize }) }
}

type Restyled = ReturnType<typeof markAttrsOf>

/** the place restyled when there is a style, and what it shows then */
function styleAt(view: EditorView, place: FieldTextPlace, attrs: Restyled | null): FieldTextStyle {
  switch (place.kind) {
    case 'text': {
      const range = restyledRange(view, place)
      if (attrs) restyle(view, range, attrs)
      return shownStyle(textElements(view, range))
    }
    case 'field':
      if (attrs) restyleFieldBlock(view, place, attrs)
      return shownStyle(fieldBlockElements(view, place))
  }
}

/** the field's place: its text, else the place an empty field or the form field index names */
function placeOf(view: EditorView, request: FieldStyleRequest) {
  const place = fieldTextPlace(view.state.doc, request)
  if (!('refused' in place)) return place
  return emptyFieldPlace(view.state.doc, request, view.state.selection.head) ?? place
}

/** vound:setFieldStyle: restyle the field's text, or refuse and leave the document as it is */
export function applyFieldStyle(view: EditorView, request: FieldStyleRequest): FieldStyleResult {
  const place = placeOf(view, request)
  if ('refused' in place) return { ok: false, reason: place.refused }
  const attrs = markAttrsOf(request.style)
  const restyled = Object.keys(attrs).length > 0 ? attrs : null
  return { ok: true, style: styleAt(view, place, restyled) }
}
