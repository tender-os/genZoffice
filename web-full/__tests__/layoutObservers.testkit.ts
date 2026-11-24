// jsdom has no layout observers: these report every observed element as on screen, and resizes only when told.
import { vi } from 'vitest'

export function stubLayoutObservers() {
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
