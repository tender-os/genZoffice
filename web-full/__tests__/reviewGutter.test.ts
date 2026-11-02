import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { chipRoom, chipRowWidth, GUTTER_RESERVE_PX, MAX_CHIP_BUTTONS, PAGE_GAP_PX, reviewFitWidth, reviewFitZoom } from '../reviewGutter'

const hostCss = readFileSync(resolve(__dirname, '../host.css'), 'utf8')

/** the scroll container's padding left and right of the page (styles.css .editor-scroll) */
const CANVAS_PAD = 24
const A4 = 793.7
const LETTER = 816

/** App.tsx at a canvas this wide in review mode: the overflow clamp compares the 100 % it starts at with
 *  rawFitWidth and refits to fitZoomFor('width'); a locked fit mode refits to fitZoomFor('width') too */
function reviewZoom(canvasWidth: number, pageWidth: number) {
  const raw = ((reviewFitWidth(canvasWidth, true) - 48) / pageWidth) * 100
  const fitted = Math.floor(reviewFitZoom(raw, true))
  return { clamped: 100 > raw + 0.5 ? fitted : 100, locked: fitted }
}
/** host.css's review grid: columns minmax(0, 1fr) | page | minmax(GUTTER_RESERVE_PX, 1fr) in the scroll
 *  container's content box. The spare width splits evenly unless the right column would get less than its
 *  minimum; it then keeps the minimum and the left column takes the rest, down to nothing. */
function pageLeft(canvasWidth: number, pageWidth: number, zoom: number) {
  const spare = canvasWidth - 2 * CANVAS_PAD - (pageWidth * zoom) / 100
  return Math.max(0, Math.min(spare / 2, spare - GUTTER_RESERVE_PX))
}
/** the gutter right of the page: from its right edge to the canvas edge */
const gutterOf = (canvasWidth: number, pageWidth: number, zoom: number) =>
  canvasWidth - CANVAS_PAD - pageLeft(canvasWidth, pageWidth, zoom) - (pageWidth * zoom) / 100

describe('review gutter', () => {
  it('reserves a three-button chip, its distance from the page and its slack, in host.css too', () => {
    expect(MAX_CHIP_BUTTONS).toBe(3)
    expect(chipRowWidth(3)).toBe(84)
    expect(GUTTER_RESERVE_PX).toBe(PAGE_GAP_PX + 84 + 16)
    expect(chipRoom(GUTTER_RESERVE_PX)).toBe(chipRowWidth(3))
    expect(hostCss).toContain(`--vound-gutter: ${GUTTER_RESERVE_PX}px;`)
  })

  it.each([860, 900, 1000, 1200, 1600, 2000, 2400])('a %ipx canvas leaves a three-button chip room right of the page', (width) => {
    for (const page of [A4, LETTER]) {
      for (const zoom of Object.values(reviewZoom(width, page))) {
        expect(zoom).toBeLessThanOrEqual(100)
        expect(pageLeft(width, page, zoom)).toBeGreaterThanOrEqual(0)
        expect(chipRoom(gutterOf(width, page, zoom))).toBeGreaterThanOrEqual(chipRowWidth(3))
      }
    }
  })

  it.each([1200, 1600, 2000])('a %ipx canvas centres the page', (width) => {
    const left = pageLeft(width, A4, 100)
    expect(CANVAS_PAD + left).toBeCloseTo(width - CANVAS_PAD - left - A4)
  })

  it('at 900px: zoomed out to the fit, the page near the canvas edge; a centred page would cover the chip', () => {
    expect(reviewZoom(900, A4)).toEqual({ clamped: 93, locked: 93 })
    expect(pageLeft(900, A4, 93)).toBeLessThan(5)
    expect(chipRoom((900 - A4 * 0.93) / 2)).toBeLessThan(chipRowWidth(3))
  })

  it('at 1000px: 100 %, moved left of the middle to keep the gutter', () => {
    expect(reviewZoom(1000, A4)).toEqual({ clamped: 100, locked: 100 })
    const left = pageLeft(1000, A4, 100)
    expect(left).toBeGreaterThan(0)
    expect(left).toBeLessThan((1000 - 2 * CANVAS_PAD - A4) / 2)
  })

  it('never fits above 100 % in review mode, and changes nothing outside it', () => {
    expect(reviewFitZoom(175, true)).toBe(100)
    expect(reviewFitZoom(80, true)).toBe(80)
    expect(reviewZoom(2400, A4).locked).toBe(100)
    expect(reviewFitZoom(175, false)).toBe(175)
    expect(reviewFitWidth(900, false)).toBe(900)
    expect(reviewFitWidth(900, true)).toBe(900 - GUTTER_RESERVE_PX)
    expect(reviewFitWidth(40, true)).toBe(0)
  })
})
