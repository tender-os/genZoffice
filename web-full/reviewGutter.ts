// Room for the field chip (fieldChip.ts) right of the page in review mode. host.css centres the page in a
// grid whose right column is never narrower than GUTTER_RESERVE_PX (--vound-gutter): on a wide canvas the
// page sits in the middle, on a narrower one it moves left until it reaches the canvas edge at the width
// fit. The app's fit-to-width zoom (App.tsx rawFitWidth and fitZoomFor, patched by the reviewFitZoom
// transform in vite.config.ts) fits the page into the canvas less GUTTER_RESERVE_PX and never above 100 %. A canvas narrower than the page plus the reserve
// therefore zooms the page out instead of covering the gutter, also when the app locked into fit mode.

/** the chip's distance from the page's right edge */
export const PAGE_GAP_PX = 10
/** the room the gutter keeps beyond the chip */
export const GUTTER_SLACK_PX = 16
/** host.css .vound-chip: buttons 2px apart and a 1px border; regular 24px buttons in 3px padding, compact 18px ones
 *  in 2px for a gutter too narrow for a row of regular ones */
const CHIP_SIZES = {
  regular: { button: 24, padding: 3 },
  compact: { button: 18, padding: 2 },
} as const
export type ChipSize = keyof typeof CHIP_SIZES
/** the sizes a chip tries, largest first */
export const CHIP_SIZE_ORDER: ChipSize[] = ['regular', 'compact']
const BUTTON_GAP_PX = 2
const BORDER_PX = 1

/** the narrowest a chip row with this many buttons can be (its label shrinks or wraps above them) */
export const chipRowWidth = (buttons: number, size: ChipSize = 'regular') =>
  buttons * CHIP_SIZES[size].button + (buttons - 1) * BUTTON_GAP_PX + 2 * (CHIP_SIZES[size].padding + BORDER_PX)

/** the most buttons the host gives one field's chip (accept, reject, ask) */
export const MAX_CHIP_BUTTONS = 3

/** what the width fit keeps free right of the page: a full chip row, its distance and its slack */
export const GUTTER_RESERVE_PX = PAGE_GAP_PX + chipRowWidth(MAX_CHIP_BUTTONS) + GUTTER_SLACK_PX
/** the highest zoom (%) a fit reaches in review mode */
export const REVIEW_MAX_FIT_ZOOM = 100

/** the canvas width (px) the app's width fit may fill: all of it, or less the chip's gutter in review mode */
export const reviewFitWidth = (canvasWidth: number, review: boolean) =>
  review ? Math.max(0, canvasWidth - GUTTER_RESERVE_PX) : canvasWidth

/** a fitted zoom (%), capped in review mode */
export const reviewFitZoom = (zoom: number, review: boolean) => (review ? Math.min(zoom, REVIEW_MAX_FIT_ZOOM) : zoom)

/** the widest a chip may be in a gutter this wide (px from the page's right edge to the canvas edge) */
export const chipRoom = (gutter: number) => gutter - PAGE_GAP_PX - GUTTER_SLACK_PX
