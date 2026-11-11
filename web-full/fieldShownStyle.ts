// What a restyled field's text shows (fieldStyle.ts): face and size as the rendered text shows them, whether the
// size is the field's own, the paragraph's alignment and the vertical alignment of the table cell around it.
import type { Node as PmNode } from '@tiptap/pm/model'
import type { EditorView } from '@tiptap/pm/view'
import { FIELD_ALIGNMENTS, type FieldAlignment, type FieldShownStyle, type FieldVerticalAlignment } from './embed-protocol'

export type Range = { from: number; to: number }

const PT_PER_PX = 0.75

// the computed text-align of a left-to-right paragraph that sets none
const CSS_ALIGNMENTS: Record<string, FieldAlignment> = { start: 'left', end: 'right', left: 'left', center: 'center', right: 'right' }

export function textSegments(doc: PmNode, { from, to }: Range) {
  const segments: Range[] = []
  doc.nodesBetween(from, to, (node, pos) => {
    if (node.isText) segments.push({ from: Math.max(pos, from), to: Math.min(pos + node.nodeSize, to) })
    return !node.isText
  })
  return segments
}

function elementAt(view: EditorView, pos: number) {
  const { node } = view.domAtPos(pos)
  return node instanceof Element ? node : node.parentElement
}

/** the elements that show the range's text, or the one around an empty range */
export function textElements(view: EditorView, range: Range) {
  const segments = textSegments(view.state.doc, range)
  const positions = segments.length > 0 ? segments.map((segment) => segment.from + 1) : [range.from]
  return positions.flatMap((pos) => {
    const element = elementAt(view, pos)
    return element ? [element] : []
  })
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

/** the face and size the elements show, each left out when they show more than one */
export function shownFaceAndSize(elements: Element[]): FieldShownStyle {
  const styles = elements.map((element) => getComputedStyle(element))
  const fontFamily = only(styles.map((style) => firstFamily(style.fontFamily)))
  const fontSize = only(styles.map((style) => pointsOf(style.fontSize)))
  return { ...(fontFamily ? { fontFamily } : {}), ...(fontSize === undefined ? {} : { fontSize }) }
}

/** the alignment a paragraph element shows; none for justified text or one the page does not lay out */
export function shownAlignment(element: Element | null): FieldAlignment | undefined {
  if (!element) return undefined
  const align = CSS_ALIGNMENTS[getComputedStyle(element).textAlign]
  return FIELD_ALIGNMENTS.find((known) => known === align)
}

/** the table cell around a document position, with its position */
export function cellAround(doc: PmNode, pos: number) {
  const $pos = doc.resolve(pos)
  for (let depth = $pos.depth; depth > 0; depth--) {
    const node = $pos.node(depth)
    if (node.type.name === 'docTableCell') return { node, pos: $pos.before(depth) }
  }
  return null
}

const VERTICAL: Record<string, FieldVerticalAlignment> = { center: 'center', bottom: 'bottom' }

/** a cell's vertical alignment: top unless it says otherwise */
export const verticalOf = (cell: PmNode): FieldVerticalAlignment => VERTICAL[String(cell.attrs.vAlign)] ?? 'top'
