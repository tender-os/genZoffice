// @vitest-environment jsdom
// The thumbnails' pages are read from the canvas' page breaks. A break inside a multi-cell table row
// has no gap, only a zero-height cut marker; pages must still be counted there, or every later page gets
// the wrong number.
import { afterEach, describe, expect, it } from 'vitest'
import { snapshotPages } from '../src/renderer/components/NavPageThumbs'

const place = (el: HTMLElement, top: number, height: number) => {
  el.getBoundingClientRect = () => new DOMRect(0, top, 800, height)
  el.getClientRects = () => [el.getBoundingClientRect()] as unknown as DOMRectList
  return el
}

const div = (className: string) => {
  const el = document.createElement('div')
  el.className = className
  return el
}

/** a page gap with no margins of its own: the paper above ends at its top, the next starts at its bottom */
const gap = (top: number) => place(div('page-gap'), top, 28)
const cut = (top: number) => place(div('page-gap-cut'), top, 0)

function canvas(...breaks: HTMLElement[]) {
  const zoom = div('doc-zoom')
  zoom.style.setProperty('--page-h', '1000')
  const wrap = place(div('page-wrap'), 0, 4000)
  const pm = place(div('doc-page ProseMirror'), 0, 4000)
  pm.append(...breaks)
  wrap.append(pm)
  zoom.append(wrap)
  document.body.append(zoom)
  return { pm, wrap }
}

const pagesOf = (pm: HTMLElement) => snapshotPages(pm, 1)?.pages

afterEach(() => {
  document.body.innerHTML = ''
})

describe('snapshotPages', () => {
  it('starts a page at a cut inside a table row', () => {
    const { pm } = canvas(gap(1000), cut(1900), gap(2500))
    expect(pagesOf(pm)).toEqual([
      { top: 0, height: 1000 },
      { top: 1028, height: 872 },
      { top: 1900, height: 600 },
      { top: 2528, height: 1000 },
    ])
  })

  it('counts a cut in a nested table, which is drawn over the page instead of in the document', () => {
    const { pm, wrap } = canvas(gap(1000), gap(2500))
    const overlays = div('page-cut-overlays')
    overlays.append(cut(1900))
    wrap.append(overlays)
    expect(pagesOf(pm)?.map((page) => page.top)).toEqual([0, 1028, 1900, 2528])
  })

  it('ignores a cut marker the stylesheet does not render', () => {
    const hidden = cut(0)
    hidden.getClientRects = () => [] as unknown as DOMRectList
    const { pm } = canvas(gap(1000), hidden)
    expect(pagesOf(pm)?.map((page) => page.top)).toEqual([0, 1028])
  })
})
