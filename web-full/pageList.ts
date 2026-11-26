// The page list left of the document in review mode: a thumbnail per page with what is still open on it, to jump
// between pages. Pages and thumbnails come from the app's own Pages tab (NavPageThumbs.tsx), so they follow the
// editor's pagination and what is typed. A press only scrolls: the caret, and with it the host's current field,
// stays. The list hides for a one-page document and in a canvas too narrow for it (host.css), and tells the host
// its width (vound:pageList), which places its own bar over the document beside it.
// The list is fixed to the frame's left edge and host.css pads the workspace by its width, so the app's own
// elements stay as they are.
import { Plugin, PluginKey } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { pageIndexAtY, snapshotPages, type ThumbPage, type ThumbSnapshot } from '../apps/docs/src/renderer/components/NavPageThumbs'
import { pageGapsOf } from '../apps/docs/src/renderer/editor/pagination-gaps'
import { embedLang } from './embedLocale'
import type { EditorToHost, SlotMarker } from './embed-protocol'
import { locateParagraphs } from './markers'
import { markersKey } from './markersPlugin'
import { countByPage, pageListLabels, pageTitle, type PageCounts, type PlacedField } from './pageCounts'

/** pagination follows an edit; one refresh for the burst */
const REFRESH_DELAY_MS = 250
/** the space a jump leaves between the host's bar and the page */
const PAGE_TOP_GAP_PX = 8
/** on the root while the list has something to show (host.css) */
const SHOWN_CLASS = 'vound-pages'
/** the marks and ids that find a field in the document; a thumbnail's copy must not answer to them */
const DOCUMENT_ONLY_ATTRIBUTES = ['id', 'contenteditable', 'data-vound-slot', 'data-src-p']
/** what the editor draws into the text that is no text of the document: page gaps with their header and footer
 *  strips, cuts, chips */
const DRAWN_WIDGET = '[contenteditable="false"]'

const isDefined = <T>(value: T | null | undefined): value is T => value !== null && value !== undefined

/** the canvas zoom the app set on .doc-zoom (1 without one) */
const zoomOf = (view: EditorView) => {
  const docZoom = view.dom.closest('.doc-zoom')
  return (docZoom && Number.parseFloat(getComputedStyle(docZoom).getPropertyValue('zoom'))) || 1
}

/** the middle of the paragraph's first line of text, in viewport px; the top of its box when it has no text.
 *  The line, not the box: a box can start just above the end of a page whose last line it does not fit on, or
 *  hold the page break (and the page's footer and header) before its text */
function firstLineMiddle(paragraph: HTMLElement) {
  const texts = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT)
  for (let text = texts.nextNode(); text; text = texts.nextNode()) {
    if (!text.textContent?.trim() || text.parentElement?.closest(DRAWN_WIDGET)) continue
    const range = document.createRange()
    range.selectNodeContents(text)
    const line = [...range.getClientRects()].find((rect) => rect.height > 0)
    if (line) return (line.top + line.bottom) / 2
  }
  return paragraph.getBoundingClientRect().top
}

/** the marked fields and where each is drawn, in unzoomed px from the top of the canvas. A field is where
 *  its paragraph's first line is, not where the marked block starts: a cell split by a page break starts on
 *  the page above the paragraph that holds the field */
function placedFields(view: EditorView, canvasTop: number, zoom: number): PlacedField[] {
  const state = markersKey.getState(view.state)
  if (!state) return []
  const marked = new Set<string>(state.set.find().flatMap((decoration) => decoration.spec.slotIds))
  const paragraphs = locateParagraphs(view.state.doc, state.markers.map((marker) => marker.paragraph))
  const fieldY = (marker: SlotMarker) => {
    const paragraph = paragraphs.get(marker.paragraph)
    const element = paragraph && view.nodeDOM(paragraph.pos)
    return element instanceof HTMLElement ? (firstLineMiddle(element) - canvasTop) / zoom : null
  }
  // text the reviewer wrote outside a field is no field of the form
  return state.markers
    .filter((marker) => marker.entry !== 'ownText' && marked.has(marker.id))
    .map((marker) => {
      const y = fieldY(marker)
      return y === null ? null : { tone: marker.tone, y }
    })
    .filter(isDefined)
}

function copyOf(block: HTMLElement) {
  const copy = block.cloneNode(true) as HTMLElement
  for (const name of DOCUMENT_ONLY_ATTRIBUTES) {
    copy.removeAttribute(name)
    for (const element of copy.querySelectorAll(`[${name}]`)) element.removeAttribute(name)
  }
  // the slot already sits where the block's margins and transforms put it
  copy.style.margin = '0'
  copy.style.transform = 'none'
  return copy
}

/** copies of the blocks on a page, placed as on the page */
function drawPage(target: HTMLElement, snapshot: ThumbSnapshot, page: ThumbPage) {
  const bottom = page.top + page.height
  const slots = snapshot.blocks
    .filter((block) => block.top < bottom && block.top + block.height > page.top)
    .map((block) => {
      const slot = document.createElement('div')
      slot.className = 'nav-thumb-slot'
      slot.style.top = `${block.top - page.top}px`
      slot.style.left = `${block.left - page.left}px`
      slot.style.width = `${block.width}px`
      slot.append(copyOf(block.el))
      return slot
    })
  target.replaceChildren(...slots)
}

function bubble(kind: 'missing' | 'review', count: number) {
  const element = document.createElement('span')
  element.className = `vound-page-bubble vound-page-bubble-${kind}`
  element.textContent = String(count)
  return element
}

type Thumb = {
  button: HTMLButtonElement
  paper: HTMLElement
  sheet: HTMLElement
  bubbles: HTMLElement
  number: HTMLElement
  onScreen: boolean
}

function createThumb(): Thumb {
  const button = document.createElement('button')
  button.type = 'button'
  button.className = 'vound-page'
  const paper = document.createElement('span')
  paper.className = 'vound-page-paper'
  const sheet = document.createElement('div')
  // the document's own classes, so the copies look as on the page
  sheet.className = 'nav-thumb-page doc-page ProseMirror'
  sheet.setAttribute('aria-hidden', 'true')
  const bubbles = document.createElement('span')
  bubbles.className = 'vound-page-bubbles'
  const number = document.createElement('span')
  number.className = 'vound-page-number'
  paper.append(sheet, bubbles)
  button.append(paper, number)
  return { button, paper, sheet, bubbles, number, onScreen: false }
}

/** shows the page list of `view` until unmounted and posts its width through `post` whenever it changes */
function mountPageList(view: EditorView, post: (msg: EditorToHost) => void) {
  const root = document.documentElement
  const labels = pageListLabels(embedLang(location.search, true))
  const list = document.createElement('nav')
  list.className = 'vound-page-list'
  list.setAttribute('aria-label', labels.list)
  const heading = document.createElement('div')
  heading.className = 'vound-page-list-heading'
  heading.textContent = labels.heading
  const body = document.createElement('div')
  body.className = 'vound-page-list-body'
  list.append(heading, body)
  document.body.append(list)

  /** the host's bar over the top of the document (host.css) */
  const barHeight = () => Number.parseFloat(getComputedStyle(root).getPropertyValue('--vound-review-top')) || 0
  const thumbs: Thumb[] = []
  let snapshot: ThumbSnapshot | null = null
  let activeIndex = -1
  let reportedWidth: number | null = null
  let scroller: HTMLElement | null = null
  let canvas: HTMLElement | null = null
  let refreshTimer: ReturnType<typeof setTimeout> | undefined
  let scrollFrame = 0

  const draw = (index: number) => {
    if (snapshot && thumbs[index]?.onScreen) drawPage(thumbs[index].sheet, snapshot, snapshot.pages[index])
  }
  const visibility = new IntersectionObserver(
    (entries) => {
      for (const entry of entries) {
        const index = thumbs.findIndex((thumb) => thumb.button === entry.target)
        if (index === -1) continue
        thumbs[index].onScreen = entry.isIntersecting
        if (entry.isIntersecting) draw(index)
        else thumbs[index].sheet.replaceChildren()
      }
    },
    { root: body, rootMargin: '200px 0px' },
  )

  /** the page at the top of the view, where a jump puts the pressed page however short it is; the last page
   *  once the view cannot scroll further, as it may never reach the top */
  const pageInView = (shown: ThumbSnapshot, scrolled: HTMLElement, drawn: HTMLElement) => {
    if (scrolled.scrollTop + scrolled.clientHeight >= scrolled.scrollHeight - 1) return shown.pages.length - 1
    const top = scrolled.getBoundingClientRect().top + barHeight() + PAGE_TOP_GAP_PX + 1
    return pageIndexAtY(shown.pages, (top - drawn.getBoundingClientRect().top) / zoomOf(view))
  }
  const markActive = () => {
    scrollFrame = 0
    if (!snapshot || !scroller || !canvas) return
    const index = pageInView(snapshot, scroller, canvas)
    if (index === activeIndex) return
    thumbs[activeIndex]?.button.removeAttribute('aria-current')
    activeIndex = index
    thumbs[index].button.setAttribute('aria-current', 'page')
    thumbs[index].button.scrollIntoView({ block: 'nearest' })
  }
  const onScroll = () => {
    if (scrollFrame === 0) scrollFrame = requestAnimationFrame(markActive)
  }

  const reportWidth = () => {
    const width = list.offsetWidth
    if (width === reportedWidth) return
    reportedWidth = width
    post({ type: 'vound:pageList', width })
  }

  const resized = new ResizeObserver(() => scheduleRefresh())
  /** follow the app's scroll container and canvas, which it may render anew */
  const follow = () => {
    const nextScroller = view.dom.closest<HTMLElement>('.editor-scroll')
    const nextCanvas = view.dom.closest<HTMLElement>('.page-wrap')
    if (nextScroller !== scroller) {
      scroller?.removeEventListener('scroll', onScroll)
      nextScroller?.addEventListener('scroll', onScroll, { passive: true })
      scroller = nextScroller
    }
    if (nextCanvas !== canvas) {
      if (canvas) resized.unobserve(canvas)
      if (nextCanvas) resized.observe(nextCanvas)
      canvas = nextCanvas
    }
  }

  const setThumbCount = (count: number) => {
    for (const thumb of thumbs.splice(count)) {
      visibility.unobserve(thumb.button)
      thumb.button.remove()
    }
    while (thumbs.length < count) {
      const thumb = createThumb()
      thumbs.push(thumb)
      body.append(thumb.button)
      visibility.observe(thumb.button)
    }
    if (activeIndex >= count) activeIndex = -1
  }

  const updateThumb = (index: number, shown: ThumbSnapshot, counts: PageCounts) => {
    const { button, paper, sheet, bubbles, number } = thumbs[index]
    const page = shown.pages[index]
    const scale = paper.clientWidth / page.width
    paper.style.height = `${page.height * scale}px`
    sheet.style.cssText = `${shown.vars};width:${page.width}px;height:${page.height}px;transform:scale(${scale})`
    const title = pageTitle(counts, labels)
    const name = `${labels.page(index + 1)}: ${title}`
    if (button.title !== title) button.title = title
    if (button.getAttribute('aria-label') !== name) button.setAttribute('aria-label', name)
    number.textContent = String(index + 1)
    bubbles.replaceChildren(
      ...(counts.missing > 0 ? [bubble('missing', counts.missing)] : []),
      ...(counts.toReview > 0 ? [bubble('review', counts.toReview)] : []),
    )
    draw(index)
  }

  const refresh = () => {
    if (view.isDestroyed) return
    follow()
    const zoom = zoomOf(view)
    snapshot = snapshotPages(view.dom, zoom)
    root.classList.toggle(SHOWN_CLASS, snapshot !== null && snapshot.pages.length > 1)
    reportWidth()
    if (!snapshot || !canvas) return
    const counts = countByPage(snapshot.pages, placedFields(view, canvas.getBoundingClientRect().top, zoom))
    setThumbCount(snapshot.pages.length)
    for (const index of thumbs.keys()) updateThumb(index, snapshot, counts[index])
    markActive()
  }
  const scheduleRefresh = () => {
    clearTimeout(refreshTimer)
    refreshTimer = setTimeout(refresh, REFRESH_DELAY_MS)
  }

  const jumpTo = (page: ThumbPage) => {
    if (!scroller || !canvas) return
    const canvasTop = canvas.getBoundingClientRect().top - scroller.getBoundingClientRect().top + scroller.scrollTop
    const reduced = window.matchMedia('(prefers-reduced-motion: reduce)').matches
    scroller.scrollTo({
      top: canvasTop + page.top * zoomOf(view) - barHeight() - PAGE_TOP_GAP_PX,
      behavior: reduced ? 'auto' : 'smooth',
    })
  }
  // a press keeps the focus, and with it the caret, in the document
  list.addEventListener('mousedown', (event) => event.preventDefault())
  body.addEventListener('click', (event) => {
    const button = event.target instanceof Element ? event.target.closest('button') : null
    const page = snapshot?.pages[thumbs.findIndex((thumb) => thumb.button === button)]
    if (page) jumpTo(page)
  })
  window.addEventListener('resize', scheduleRefresh)
  scheduleRefresh()

  return {
    scheduleRefresh,
    unmount: () => {
      clearTimeout(refreshTimer)
      cancelAnimationFrame(scrollFrame)
      window.removeEventListener('resize', scheduleRefresh)
      scroller?.removeEventListener('scroll', onScroll)
      resized.disconnect()
      visibility.disconnect()
      root.classList.remove(SHOWN_CLASS)
      list.remove()
    },
  }
}

export const pageListKey = new PluginKey('voundPageList')

/** the page list of the editor in review mode (reviewMode.ts) */
export const pageListPlugin = (post: (msg: EditorToHost) => void) =>
  new Plugin({
    key: pageListKey,
    view: (view) => {
      const pages = mountPageList(view, post)
      return {
        update: (updated, before) => {
          const { state } = updated
          const marksChanged = markersKey.getState(state)?.set !== markersKey.getState(before)?.set
          const pagesChanged = pageGapsOf(state) !== pageGapsOf(before)
          if (marksChanged || pagesChanged || state.doc !== before.doc) pages.scheduleRefresh()
        },
        destroy: pages.unmount,
      }
    },
  })
