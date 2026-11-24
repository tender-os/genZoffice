// @vitest-environment jsdom
import { EditorState, TextSelection } from '@tiptap/pm/state'
import { EditorView } from '@tiptap/pm/view'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { EditorToHost, SlotMarker } from '../embed-protocol'
import { markersPlugin, setMarkers } from '../markersPlugin'
import { pageListPlugin } from '../pageList'
import { stubLayoutObservers } from './layoutObservers.testkit'
import { schema } from './markerSchema'

const LINE = 100

/** gives an element a box at `top`, 800 wide */
const place = (element: Element, top: number, height: number) => {
  element.getBoundingClientRect = () => new DOMRect(0, top, 800, height)
}

const marker = (id: string, paragraph: number, tone: SlotMarker['tone']): SlotMarker => ({
  id, paragraph, tone, scope: 'paragraph', label: id, actions: ['ask'],
})

/** four paragraphs of 100px with page breaks above those named in `breaksBefore` */
function mount(breaksBefore: number[]) {
  document.body.innerHTML =
    '<main class="editor-scroll"><div class="doc-zoom"><div class="page-wrap"><div class="host"></div></div></div></main>'
  const scroller = document.querySelector<HTMLElement>('.editor-scroll')!
  const scrollTo = vi.fn()
  scroller.scrollTo = scrollTo
  const wrap = document.querySelector<HTMLElement>('.page-wrap')!
  place(scroller, 0, 500)
  place(wrap, 0, 4 * LINE)
  const posted: EditorToHost[] = []
  const paragraphs = [0, 1, 2, 3].map((i) => schema.nodes.docParagraph.create({ docxPara: i }, schema.text(`text ${i}`)))
  const view = new EditorView(document.querySelector('.host')!, {
    state: EditorState.create({ doc: schema.nodes.doc.create(null, paragraphs), plugins: [markersPlugin, pageListPlugin((msg) => posted.push(msg))] }),
  })
  const layOut = () => {
    Array.from(view.dom.children).forEach((block, i) => place(block, i * LINE, LINE))
  }
  layOut()
  for (const index of breaksBefore) {
    const gap = document.createElement('div')
    gap.className = 'page-gap'
    place(gap, index * LINE - 10, 10)
    wrap.append(gap)
  }
  return {
    view, posted, scrollTo,
    settle: () => {
      layOut()
      vi.advanceTimersByTime(300)
    },
  }
}

const pages = () => [...document.querySelectorAll<HTMLElement>('.vound-page')]

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

describe('the page list', () => {
  it('shows what is open on each page', () => {
    const { view, settle } = mount([2])
    setMarkers(view, [marker('a', 0, 'accepted'), marker('b', 2, 'gap'), marker('c', 3, 'proposed')])
    settle()
    expect(document.documentElement.classList.contains('vound-pages')).toBe(true)
    expect(pages().map((page) => page.getAttribute('aria-label'))).toEqual([
      'Page 1: Fully filled in',
      'Page 2: 1 missing · 1 to review',
    ])
  })

  it('scrolls to a pressed page and leaves the caret where it is', () => {
    const { view, posted, scrollTo, settle } = mount([2])
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 3)))
    settle()
    const before = posted.length
    pages()[1].click()
    // page 2 starts where the gap ends, less the 8px kept above a page
    expect(scrollTo).toHaveBeenCalledWith({ top: 2 * LINE - 8, behavior: 'auto' })
    expect(view.state.selection.from).toBe(3)
    expect(posted.length).toBe(before)
  })

  it('copies the page without what finds a field in the document', () => {
    const { view, settle } = mount([2])
    setMarkers(view, [marker('a', 0, 'gap')])
    settle()
    expect(pages()[0].textContent).toContain('text 0')
    expect(pages()[0].querySelector('.vound-slot-gap')).not.toBeNull()
    expect(document.querySelectorAll('[data-vound-slot]')).toHaveLength(1)
  })

  it('is not shown for a document of one page', () => {
    const { settle } = mount([])
    settle()
    expect(document.documentElement.classList.contains('vound-pages')).toBe(false)
  })

  it('goes with the editor', () => {
    const { view, settle } = mount([2])
    settle()
    view.destroy()
    expect(document.querySelector('.vound-page-list')).toBeNull()
    expect(document.documentElement.classList.contains('vound-pages')).toBe(false)
  })
})
