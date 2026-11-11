// The host's Format control (vound:setFieldStyle): the face and size of a field's text, its paragraph's alignment and
// the vertical alignment of the table cell around it. The target names the field as a write does (fieldWrite.ts),
// else as an empty field (fieldStylePlace.ts); the selection is restyled when it lies inside the field, else the
// whole field, and a field passthrough's display (fieldStyleBlock.ts). Every restyle is one undoable transaction with
// HOST_WRITE_META, saved as the run's w:rFonts and w:sz, the paragraph's w:jc and the cell's w:vAlign; an empty place
// stores the face and size for the text typed there next. The answer is what the text shows now
// (fieldShownStyle.ts); an empty style changes nothing and only answers.
import type { Node as PmNode } from '@tiptap/pm/model'
import { TextSelection, type Transaction } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import type { FieldShownStyle, FieldTextStyle, FieldValueRefusal } from './embed-protocol'
import { cellAround, shownAlignment, shownFaceAndSize, textElements, textSegments, verticalOf, type Range } from './fieldShownStyle'
import { restyleFieldBlock, shownFieldBlock } from './fieldStyleBlock'
import { emptyFieldPlace } from './fieldStylePlace'
import { fieldTextPlace, HOST_WRITE_META, type FieldTarget, type FieldTextPlace } from './fieldWrite'

/** `formText`: the index of the legacy text form field in its paragraph, for a field its text does not name once */
export type FieldStyleRequest = FieldTarget & { formText?: number; style: FieldTextStyle }
export type FieldStyleResult = { ok: true; style: FieldShownStyle } | { ok: false; reason: FieldValueRefusal }

const halfPointsOf = (points: number | null) => (points === null ? null : Math.round(points * 2))

/** the docTextStyle attributes of a style: the Latin face, and the size in half-points or null for none of its own */
export const markAttrsOf = (style: FieldTextStyle) => ({
  ...(style.fontFamily === undefined ? {} : { fontAscii: style.fontFamily }),
  ...(style.fontSize === undefined ? {} : { sizeHalfPoints: halfPointsOf(style.fontSize) }),
})

/** the selection when it lies inside the field, else the field */
function restyledRange(view: EditorView, field: Range): Range {
  const { from, to, empty } = view.state.selection
  return !empty && from >= field.from && to <= field.to ? { from, to } : field
}

/** the paragraphs the range lies in, or the one an empty range sits in */
function paragraphsOf(doc: PmNode, range: Range) {
  const paragraphs: { node: PmNode; pos: number }[] = []
  doc.nodesBetween(range.from, range.to, (node, pos) => {
    if (node.isTextblock) paragraphs.push({ node, pos })
    return !node.isTextblock
  })
  if (paragraphs.length > 0) return paragraphs
  const $from = doc.resolve(range.from)
  return $from.parent.isTextblock ? [{ node: $from.parent, pos: $from.before() }] : []
}

function withMarks(tr: Transaction, range: Range, attrs: ReturnType<typeof markAttrsOf>) {
  const type = tr.doc.type.schema.marks.docTextStyle
  if (range.from === range.to) {
    const marks = tr.doc.resolve(range.from).marks()
    const existing = type.isInSet(marks)
    const stored = [...marks.filter((mark) => mark.type !== type), type.create({ ...existing?.attrs, ...attrs })]
    return tr.setSelection(TextSelection.create(tr.doc, range.from)).setStoredMarks(stored)
  }
  for (const segment of textSegments(tr.doc, range)) {
    const existing = type.isInSet(tr.doc.resolve(segment.from + 1).marks())
    tr.addMark(segment.from, segment.to, type.create({ ...existing?.attrs, ...attrs }))
  }
  return tr
}

function restyleText(view: EditorView, range: Range, style: FieldTextStyle) {
  const { doc } = view.state
  const attrs = markAttrsOf(style)
  const tr = Object.keys(attrs).length > 0 ? withMarks(view.state.tr, range, attrs) : view.state.tr
  const { alignment, verticalAlignment } = style
  if (alignment) for (const { node, pos } of paragraphsOf(doc, range)) tr.setNodeMarkup(pos, undefined, { ...node.attrs, align: alignment })
  const cell = verticalAlignment ? cellAround(doc, range.from) : null
  if (cell) tr.setNodeMarkup(cell.pos, undefined, { ...cell.node.attrs, vAlign: verticalAlignment, tcPrEdited: true })
  if (!tr.docChanged && !tr.storedMarksSet) return
  view.dispatch(tr.setMeta(HOST_WRITE_META, true))
  if (range.from === range.to) view.focus()
}

/** whether the range's text sets its own size, every part of it */
function ownSize(view: EditorView, range: Range) {
  const { doc, storedMarks } = view.state
  const sizeOf = (marks: readonly { type: { name: string }; attrs: Record<string, unknown> }[]) =>
    marks.find((mark) => mark.type.name === 'docTextStyle')?.attrs.sizeHalfPoints
  const segments = textSegments(doc, range)
  if (segments.length === 0) return sizeOf(storedMarks ?? doc.resolve(range.from).marks()) != null
  return segments.every((segment) => sizeOf(doc.resolve(segment.from + 1).marks()) != null)
}

function shownText(view: EditorView, range: Range): FieldShownStyle {
  const { doc } = view.state
  const paragraph = paragraphsOf(doc, range)[0]
  const dom = paragraph ? view.nodeDOM(paragraph.pos) : null
  const alignment = shownAlignment(dom instanceof Element ? dom : null)
  const cell = cellAround(doc, range.from)
  return {
    ...shownFaceAndSize(textElements(view, range)),
    sizing: ownSize(view, range) ? 'own' : 'document',
    ...(alignment ? { alignment } : {}),
    ...(cell ? { verticalAlignment: verticalOf(cell.node) } : {}),
  }
}

/** the place restyled, and what it shows then */
function styleAt(view: EditorView, place: FieldTextPlace, style: FieldTextStyle): FieldShownStyle {
  switch (place.kind) {
    case 'text': {
      const range = restyledRange(view, place)
      restyleText(view, range, style)
      return shownText(view, range)
    }
    case 'field':
      restyleFieldBlock(view, place, style)
      return shownFieldBlock(view, place)
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
  return { ok: true, style: styleAt(view, place, request.style) }
}
