// jsdom has no layout observers: these report every observed element as on screen, and resizes only when told.
// It draws no text either.
import { vi } from 'vitest'

export function stubLayoutObservers() {
  // nor does it draw text: a range has no line boxes (a test that needs lines sets its own)
  Range.prototype.getClientRects = () => [] as unknown as DOMRectList
  vi.stubGlobal(
    'IntersectionObserver',
    class {
      constructor(private report: (entries: Array<{ target: Element; isIntersecting: boolean }>) => void) {}
      observe(target: Element) {
        this.report([{ target, isIntersecting: true }])
      }
      unobserve() {}
      disconnect() {}
    },
  )
  vi.stubGlobal(
    'ResizeObserver',
    class {
      observe() {}
      unobserve() {}
      disconnect() {}
    },
  )
}
