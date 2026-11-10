// A restyle of a field passthrough paragraph (fieldStyle.ts): its result text is no editor text, so the face and size
// go into its display. The runs of the restyled part take them, and the display's own face and size too when the
// part is all of the text or empty, so text typed into the field later keeps them (the editor's field commit drops
// the formatted runs). The save writes them into the paragraph's runs (convert.ts, passthroughFaceAndSize).
import type { FieldRun } from '@genoffice/docx-engine'
import type { EditorView } from '@tiptap/pm/view'
import { FIELD_TEXT_SELECTOR, HOST_WRITE_META } from './fieldWrite'

type FaceAndSize = { fontAscii?: string; sizeHalfPoints?: number }
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

const displayFaceAndSize = (style: FaceAndSize) => ({
  ...(style.fontAscii === undefined ? {} : { fontFamily: style.fontAscii }),
  ...(style.sizeHalfPoints === undefined ? {} : { szHalfPoints: style.sizeHalfPoints }),
})

/** the display's runs with the part restyled; an empty part has no text to restyle */
function restyledRuns(display: { left?: string; runs?: FieldRun[] }, place: Place, style: FaceAndSize) {
  if (place.from === place.to) return display.runs
  const runs = display.runs ?? [{ text: display.left ?? '' }]
  return cutRuns(runs, place.from, place.to).map(({ run, inside }) => (inside ? { ...run, ...style } : run))
}

export function restyleFieldBlock(view: EditorView, place: Place, style: FaceAndSize) {
  const node = view.state.doc.nodeAt(place.pos)
  if (!node) return
  const display = node.attrs.fieldDisplay
  const length = (display.left ?? '').length
  const spansAll = place.from === place.to || (place.from === 0 && place.to === length)
  const runs = restyledRuns(display, place, style)
  const next = { ...display, ...(runs ? { runs } : {}), ...(spansAll ? displayFaceAndSize(style) : {}) }
  view.dispatch(view.state.tr.setNodeMarkup(place.pos, undefined, { ...node.attrs, fieldDisplay: next }).setMeta(HOST_WRITE_META, true))
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
