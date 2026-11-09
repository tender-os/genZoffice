import { describe, expect, it } from 'vitest'
import { zoomRefit } from '../src/renderer/zoom-refit'

describe('zoomRefit', () => {
  it('follows the fit mode while the zoom is the last fit', () => {
    expect(
      zoomRefit({ zoom: 53, raw: 60, previousRaw: 53, lastFit: { mode: 'page', value: 53 } }),
    ).toBe('page')
  })

  it('clamps a zoom the page no longer fits at when the width fit changed', () => {
    expect(zoomRefit({ zoom: 80, raw: 53, previousRaw: 90, lastFit: null })).toBe('width')
  })

  it('clamps an overflowing zoom on the first check', () => {
    expect(zoomRefit({ zoom: 100, raw: 53, previousRaw: null, lastFit: null })).toBe('width')
  })

  it('keeps a zoomed-in page when the width fit did not change, as after a save re-parse', () => {
    expect(
      zoomRefit({ zoom: 80, raw: 53, previousRaw: 53, lastFit: { mode: 'width', value: 53 } }),
    ).toBeNull()
  })

  it('keeps a manual zoom the page still fits at', () => {
    expect(zoomRefit({ zoom: 50, raw: 53, previousRaw: 90, lastFit: null })).toBeNull()
  })
})
