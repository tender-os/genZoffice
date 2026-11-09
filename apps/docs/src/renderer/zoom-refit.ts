import type { ZoomFitMode } from './wheel-zoom'

type ZoomRefitInput = {
  zoom: number
  /** the uncapped width-fit zoom now */
  raw: number
  /** the uncapped width-fit zoom at the last check; null before the first */
  previousRaw: number | null
  lastFit: { mode: ZoomFitMode; value: number } | null
}

/** The fit to apply after the canvas or page may have changed: follow the user's fit mode, clamp a manual
 *  zoom the page no longer fits at since the width fit changed, or null to keep the zoom. */
export function zoomRefit({ zoom, raw, previousRaw, lastFit }: ZoomRefitInput): ZoomFitMode | null {
  if (lastFit !== null && Math.abs(zoom - lastFit.value) <= 0.5) return lastFit.mode
  if (raw === previousRaw || zoom <= raw + 0.5) return null
  return 'width'
}
