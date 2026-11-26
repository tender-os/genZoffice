// @vitest-environment jsdom
// The page list against layouts the plain four-paragraph canvas of pageList.test.ts does not have: a field
// whose marked cell starts a page before it, a page break drawn at the start of a paragraph, a short page, the
// end of the document, and a landscape page among portrait ones. Boxes are placed by hand, as the canvas would.
import { Plugin } from '@tiptap/pm/state'
import { Decoration, DecorationSet, EditorView } from '@tiptap/pm/view'
import { EditorState } from '@tiptap/pm/state'
import type { Node as PmNode } from '@tiptap/pm/model'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { SlotMarker } from '../embed-protocol'
import { markersPlugin, setMarkers } from '../markersPlugin'
import { pageListPlugin } from '../pageList'
import { stubLayoutObservers } from './layoutObservers.testkit'
import { cell, doc, para, paragraphs, table } from './markerSchema'

const place = (element: Element, top: number, height: number, width = 800, left = 0) => {
  element.getBoundingClientRect = () => new DOMRect(left, top, width, height)
  element.getClientRects = () => [new DOMRect(left, top, width, height)] as unknown as DOMRectList
}

const marker = (id: string, paragraph: number, scope: SlotMarker['scope'] = 'paragraph'): SlotMarker => ({
  id, paragraph, scope, tone: 'gap', label: id, actions: ['ask'],
})

const pageGap = (page?: { left: number; width: number }) => {
  const gap = document.createElement('div')
  gap.className = 'page-gap'
  if (page) {
    gap.style.setProperty('--gap-page-x', `${page.left}px`)
    gap.style.setProperty('--gap-page-w', `${page.width}px`)
  }
  return gap
}

/** a page gap drawn inside the paragraph at `pos`, before its text */
const leadingGapAt = (pos: number) =>
  new Plugin({
    props: {
      decorations: (state) =>
        DecorationSet.create(state.doc, [Decoration.widget(pos + 1, () => pageGap(), { side: -1 })]),
    },
  })

/** the editor over `content`, `layOut` placing its boxes; the view is scrolled `scrollTop` of `scrollHeight` */
function mount(content: PmNode, layOut: (view: EditorView, wrap: HTMLElement) => void, plugins: Plugin[] = []) {
  document.body.innerHTML =
    '<main class="editor-scroll"><div class="doc-zoom"><div class="page-wrap"><div class="host"></div></div></div></main>'
  const scroller = document.querySelector<HTMLElement>('.editor-scroll')!
  scroller.scrollTo = vi.fn()
  place(scroller, 0, 500)
  const scroll = (scrollTop: number, scrollHeight = 2000) => {
    Object.defineProperties(scroller, {
      scrollTop: { value: scrollTop, configurable: true },
      scrollHeight: { value: scrollHeight, configurable: true },
      clientHeight: { value: 500, configurable: true },
    })
  }
  scroll(0)
  const wrap = document.querySelector<HTMLElement>('.page-wrap')!
  const view = new EditorView(document.querySelector('.host')!, {
    state: EditorState.create({ doc: content, plugins: [markersPlugin, pageListPlugin(() => {}), ...plugins] }),
  })
  const settle = () => {
    layOut(view, wrap)
    vi.advanceTimersByTime(300)
  }
  return { view, wrap, scroller, scroll, settle }
}

const pages = () => [...document.querySelectorAll<HTMLElement>('.vound-page')]
const labels = () => pages().map((page) => page.getAttribute('aria-label'))
const marked = () => pages().findIndex((page) => page.getAttribute('aria-current') === 'page')

beforeEach(() => {
  vi.useFakeTimers()
  stubLayoutObservers()
  vi.stubGlobal('matchMedia', () => ({ matches: true }))
  Element.prototype.scrollIntoView = () => {}
})
afterEach(() => {
  vi.useRealTimers()
  vi.unstubAllGlobals()
  document.body.innerHTML = ''
  document.documentElement.className = ''
})

describe('the page a field is counted on', () => {
  it('is where its paragraph is drawn, when the cell around it starts a page earlier', () => {
    // one cell: a caption on page 1, the field's paragraph pushed below a break inside the cell
    const { view, settle } = mount(doc(table(cell(para(0, 'Name'), para(1, '____'))), para(2, 'after')), (v, wrap) => {
      const [tableBox, after] = Array.from(v.dom.children)
      place(tableBox, 0, 400)
      place(v.dom.querySelector('td')!, 0, 400)
      const [caption, field] = Array.from(v.dom.querySelectorAll('td p'))
      place(caption, 0, 100)
      place(field, 300, 100)
      place(after, 400, 100)
      place(v.dom, 0, 500)
      place(wrap, 0, 500)
      if (!wrap.querySelector(':scope > .page-gap')) wrap.append(pageGap())
      place(wrap.querySelector(':scope > .page-gap')!, 200, 10)
    })
    setMarkers(view, [marker('a', 1, 'cell')])
    settle()
    expect(labels()).toEqual(['Page 1: No input needed', 'Page 2: 1 missing'])
  })

  it('is the page after a break drawn at the start of its paragraph', () => {
    // paragraph 2's box starts at 200, but the break inside it moves its text to the next page
    const content = doc(...paragraphs(0, 4))
    let start = 0
    content.forEach((_, offset, index) => { if (index === 2) start = offset })
    const { view, settle } = mount(
      content,
      (v, wrap) => {
        Array.from(v.dom.children).forEach((block, i) => place(block, i * 100, 100))
        place(v.dom, 0, 400)
        place(wrap, 0, 400)
        place(v.dom.querySelector('p .page-gap')!, 200, 30)
      },
      [leadingGapAt(start)],
    )
    setMarkers(view, [marker('a', 2)])
    settle()
    expect(labels()).toEqual(['Page 1: No input needed', 'Page 2: 1 missing'])
  })
})

describe('the page marked in the list', () => {
  /** pages at 0, 200 (a short one, 100 tall) and 300; the canvas scrolled up by `scrolled` */
  const threePages = () =>
    mount(doc(...paragraphs(0, 4)), (v, wrap) => {
      Array.from(v.dom.children).forEach((block, i) => place(block, i * 100, 100))
      place(v.dom, 0, 400)
      place(wrap, 0, 400)
      if (!wrap.querySelector(':scope > .page-gap')) wrap.append(pageGap(), pageGap())
      const [first, second] = Array.from(wrap.querySelectorAll(':scope > .page-gap'))
      place(first, 190, 10)
      place(second, 290, 10)
    })
  const scrollCanvasTo = (wrap: HTMLElement, scroller: HTMLElement, y: number) => {
    place(wrap, -y, 400)
    scroller.dispatchEvent(new Event('scroll'))
    vi.advanceTimersByTime(20)
  }

  it('is the page at the top of the view, so a short page scrolled to the top is the one marked', () => {
    const { wrap, scroller, scroll, settle } = threePages()
    settle()
    scroll(200)
    scrollCanvasTo(wrap, scroller, 200)
    expect(marked()).toBe(1)
  })

  it('is the last page once the view is scrolled to the end', () => {
    const { wrap, scroller, scroll, settle } = threePages()
    settle()
    scroll(1500)
    scrollCanvasTo(wrap, scroller, 100)
    expect(marked()).toBe(2)
  })
})

describe('the page thumbnails', () => {
  it('draw each page at its own width, a landscape page wider than the portrait one before it', () => {
    const { settle } = mount(doc(...paragraphs(0, 4)), (v, wrap) => {
      Array.from(v.dom.children).forEach((block, i) => place(block, i * 100, 100))
      place(v.dom, 0, 400, 800)
      place(wrap, 0, 400, 1100)
      if (!wrap.querySelector(':scope > .page-gap')) wrap.append(pageGap({ left: 0, width: 1100 }))
      place(wrap.querySelector(':scope > .page-gap')!, 190, 10, 1100)
    })
    settle()
    const widths = [...document.querySelectorAll<HTMLElement>('.vound-page .nav-thumb-page')].map((sheet) => sheet.style.width)
    expect(widths).toEqual(['800px', '1100px'])
  })
})
