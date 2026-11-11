// A restyle of a field passthrough paragraph (fieldStyle.ts): its result text is no editor text, so the face and size
// go into its display. The runs of the restyled part take them, and the display's own face and size too when the
// part is all of the text or empty, so text typed into the field later keeps them; a size of null takes the size
// away. The alignment is the display's (fieldDisplay.align); a passthrough is never in a table cell. The save writes
// them into the paragraph (convert.ts, passthroughFaceAndSize).
import type { FieldRun } from '@genoffice/docx-engine'
import type { EditorView } from '@tiptap/pm/view'
import type { FieldShownStyle, FieldTextStyle } from './embed-protocol'
import { shownAlignment, shownFaceAndSize } from './fieldShownStyle'
import { FIELD_TEXT_SELECTOR, HOST_WRITE_META } from './fieldWrite'

type FaceAndSize = { fontAscii?: string; sizeHalfPoints?: number | null }
type Place = { pos: number; from: number; to: number }

/** the runs with offsets from..to of their joined text cut out as runs of their own, each marked inside or not */
function cutRuns(runs: FieldRun[], from: number, to: number) {
  let start = 0
  return runs.flatMap((run) => {
    const end = start + run.text.length
    const cuts = [start, Math.min(Math.max(from, start), end), Math.min(Math.max(to, start), end), end]
    start = end
    return [0, 1, 2].flatMap((i) => {
      const text = run.text.slice(cuts[i] - cuts[0], cuts[i + 1] - cuts[0])
      return text === '' ? [] : [{ run: { ...run, text }, inside: i === 1 }]
    })
  })
}

/** a run or display with the face, and the size or none of its own */
function withFaceAndSize<T extends { sizeHalfPoints?: number }>(run: T, style: FaceAndSize): T {
  const { sizeHalfPoints, ...face } = style
  if (sizeHalfPoints === undefined) return { ...run, ...face }
  if (sizeHalfPoints !== null) return { ...run, ...face, sizeHalfPoints }
  const { sizeHalfPoints: _dropped, ...unsized } = run
  return { ...unsized, ...face } as T
}

function displayFaceAndSize<T extends { fontFamily?: string; szHalfPoints?: number }>(display: T, style: FaceAndSize): T {
  const faced = style.fontAscii === undefined ? display : { ...display, fontFamily: style.fontAscii }
  if (style.sizeHalfPoints === undefined) return faced
  if (style.sizeHalfPoints !== null) return { ...faced, szHalfPoints: style.sizeHalfPoints }
  const { szHalfPoints: _dropped, ...unsized } = faced
  return unsized as T
}

/** the display's runs with the part restyled; an empty part has no text to restyle */
function restyledRuns(display: { left?: string; runs?: FieldRun[] }, place: Place, style: FaceAndSize) {
  if (place.from === place.to) return display.runs
  const runs = display.runs ?? [{ text: display.left ?? '' }]
  return cutRuns(runs, place.from, place.to).map(({ run, inside }) => (inside ? withFaceAndSize(run, style) : run))
}

const faceAndSizeOf = (style: FieldTextStyle): FaceAndSize => ({
  ...(style.fontFamily === undefined ? {} : { fontAscii: style.fontFamily }),
  ...(style.fontSize === undefined ? {} : { sizeHalfPoints: style.fontSize === null ? null : Math.round(style.fontSize * 2) }),
})

export function restyleFieldBlock(view: EditorView, place: Place, style: FieldTextStyle) {
  const node = view.state.doc.nodeAt(place.pos)
  const faceAndSize = faceAndSizeOf(style)
  if (!node || (Object.keys(faceAndSize).length === 0 && !style.alignment)) return
  const display = node.attrs.fieldDisplay
  const length = (display.left ?? '').length
  const spansAll = place.from === place.to || (place.from === 0 && place.to === length)
  const runs = restyledRuns(display, place, faceAndSize)
  const restyled = { ...display, ...(runs ? { runs } : {}), ...(style.alignment ? { align: style.alignment } : {}) }
  const next = spansAll ? displayFaceAndSize(restyled, faceAndSize) : restyled
  view.dispatch(view.state.tr.setNodeMarkup(place.pos, undefined, { ...node.attrs, fieldDisplay: next }).setMeta(HOST_WRITE_META, true))
}

/** whether the part sets its own size: its runs do, or the display does for a part without runs */
function ownSize(display: { runs?: FieldRun[]; szHalfPoints?: number }, place: Place) {
  if (!display.runs || place.from === place.to) return display.szHalfPoints !== undefined
  return cutRuns(display.runs, place.from, place.to).every(({ run, inside }) => !inside || run.sizeHalfPoints !== undefined)
}

export function shownFieldBlock(view: EditorView, place: Place): FieldShownStyle {
  const display = view.state.doc.nodeAt(place.pos)?.attrs.fieldDisplay ?? {}
  const box = fieldBlockElements(view, place)
  const alignment = shownAlignment(box[0]?.closest(FIELD_TEXT_SELECTOR) ?? null)
  return {
    ...shownFaceAndSize(box),
    sizing: ownSize(display, place) ? 'own' : 'document',
    ...(alignment ? { alignment } : {}),
  }
}

/** the elements that show offsets from..to of a field passthrough's result text: its spans, or the text box itself */
export function fieldBlockElements(view: EditorView, place: Place): Element[] {
  const dom = view.nodeDOM(place.pos)
  const box = dom instanceof Element ? dom.querySelector(FIELD_TEXT_SELECTOR) : null
  if (!box) return []
  let start = 0
  const spans = [...box.childNodes].flatMap((child) => {
    const end = start + (child.textContent ?? '').length
    const overlaps = child instanceof Element && start < place.to && end > place.from
    start = end
    return overlaps ? [child] : []
  })
  return spans.length > 0 ? spans : [box]
}
