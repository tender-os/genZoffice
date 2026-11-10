// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { EditorState } from '@tiptap/pm/state'
import { EditorView } from '@tiptap/pm/view'
import type { EditorToHost, FieldAction, SlotMarker } from '../embed-protocol'
import { mountFieldChip } from '../fieldChip'
import { markersKey, markersPlugin, setMarkers } from '../markersPlugin'
import { cell, doc, para, paragraphs, schema, table } from './markerSchema'

const marker = (id: string, paragraph: number, actions: FieldAction[], label = id): SlotMarker =>
  ({ id, paragraph, scope: 'paragraph', tone: 'proposed', label, actions })

type Box = { left: number; top: number; right: number; bottom: number }
const rectOf = ({ left, top, right, bottom }: Box) =>
  ({ x: left, y: top, left, top, right, bottom, width: right - left, height: bottom - top, toJSON: () => ({}) })

let view: EditorView | null = null
let unmount: (() => void) | null = null
let scroller: HTMLElement
let gutterRight = 1200

beforeEach(() => {
  gutterRight = 1200
  scroller = document.createElement('main')
  scroller.className = 'editor-scroll'
  document.body.append(scroller)
  // jsdom lays nothing out: the scroll container, the page (the editor root), the marked blocks and the
  // chip get fixed boxes, the page's right edge at 900 and 300px of gutter to its right
  vi.spyOn(Element.prototype, 'getBoundingClientRect').mockImplementation(function (this: Element) {
    if (this === scroller) return rectOf({ left: 0, top: 0, right: gutterRight, bottom: 800 })
    if (this === view?.dom) return rectOf({ left: 100, top: 0, right: 900, bottom: 2000 })
    if (this.classList.contains('vound-slot')) return rectOf({ left: 196, top: 100, right: 804, bottom: 120 })
    if (this.classList.contains('vound-chip')) return rectOf({ left: 0, top: 0, right: 60, bottom: 30 })
    return rectOf({ left: 0, top: 0, right: 0, bottom: 0 })
  })
})

afterEach(() => {
  unmount?.(); unmount = null
  view?.destroy(); view = null
  scroller.remove()
  vi.restoreAllMocks()
  vi.useRealTimers()
  document.getSelection()?.removeAllRanges()
})

/** paragraphs 0..4 in the scroll container, the chip mounted, these markers set */
function chipView(markers: SlotMarker[]) {
  const post = vi.fn<(msg: EditorToHost) => void>()
  const mount = document.createElement('div')
  scroller.append(mount)
  view = new EditorView(mount, { state: EditorState.create({ schema, doc: doc(...paragraphs(0, 5)), plugins: [markersPlugin] }) })
  setMarkers(view, markers)
  unmount = mountFieldChip(view, post)
  return { v: view, post }
}
const block = (v: EditorView, id: string) => v.dom.querySelector(`[data-vound-slot~="${id}"]`)!
const hover = (target: Element, at: MouseEventInit = {}) => target.dispatchEvent(new MouseEvent('mouseover', { bubbles: true, ...at }))
const chip = () => scroller.querySelector<HTMLElement>('.vound-chip')
const shown = () => chip() !== null && !chip()!.hidden
const buttons = () => [...chip()!.querySelectorAll('button')].map((b) => ({ text: b.textContent, title: b.title, label: b.getAttribute('aria-label') }))
const rowLabels = () => [...chip()!.querySelectorAll('.vound-chip-label')].map((l) => l.textContent)

describe('mountFieldChip', () => {
  it('hovering a marked paragraph shows its actions', () => {
    const { v } = chipView([marker('name', 2, ['accept', 'reject'], 'Name')])
    hover(block(v, 'name'))

    expect(shown()).toBe(true)
    expect(buttons()).toEqual([
      { text: '✓', title: 'Übernehmen', label: 'Übernehmen' },
      { text: '✕', title: 'Ablehnen', label: 'Ablehnen' },
    ])
    // one field: no label
    expect(rowLabels()).toEqual([])
    // in the gutter 10px right of the page, vertically centred on the block, the tether across the margin
    expect(chip()!.style.left).toBe('910px')
    expect(chip()!.style.top).toBe('110px')
    // never wider than the gutter: 300px less the 10px distance and the 16px slack
    expect(chip()!.style.maxWidth).toBe('274px')
    const tether = scroller.querySelector<HTMLElement>('.vound-chip-tether')!
    expect(tether.hidden).toBe(false)
    expect(tether.style.left).toBe('804px')
    expect(tether.style.width).toBe('106px')
    expect(tether.style.top).toBe('110px')
    // the hover outline, as a decoration
    expect(block(v, 'name').classList.contains('vound-slot-hot')).toBe(true)
  })

  it("shows no chip on the reviewer's own text", () => {
    const { v } = chipView([{ ...marker('own', 2, ['ask'], 'Eigener Text'), entry: 'ownText' }])
    hover(block(v, 'own'))
    expect(shown()).toBe(false)
  })

  it('a two-slot paragraph shows two labelled rows', () => {
    const { v } = chipView([marker('ort', 1, ['fill', 'ask'], 'Ort'), marker('datum', 1, ['restore', 'fill'], 'Datum')])
    hover(block(v, 'ort'))

    expect(shown()).toBe(true)
    expect(chip()!.querySelectorAll('.vound-chip-row')).toHaveLength(2)
    expect(rowLabels()).toEqual(['Ort', 'Datum'])
    // a label cut short in a narrow gutter can still be read
    expect([...chip()!.querySelectorAll<HTMLElement>('.vound-chip-label')].map((l) => l.title)).toEqual(['Ort', 'Datum'])
    expect(buttons().map((b) => b.title)).toEqual(['Wert eintragen', '', 'KI-Wert wiederherstellen', 'Wert eintragen'])
    expect(buttons().map((b) => b.text)).toEqual(['✎', '', '↺', '✎'])
  })

  it('clicking ✕ posts fieldAction reject with the slot id', () => {
    const { v, post } = chipView([marker('ort', 1, ['accept', 'reject'], 'Ort'), marker('datum', 1, ['accept', 'reject'], 'Datum')])
    hover(block(v, 'datum'))
    const reject = chip()!.querySelectorAll<HTMLButtonElement>('.vound-chip-row')[1].querySelector<HTMLButtonElement>('[aria-label="Ablehnen"]')!
    // a press keeps the focus (and the caret) in the document
    const press = new MouseEvent('mousedown', { bubbles: true, cancelable: true })
    reject.dispatchEvent(press)
    expect(press.defaultPrevented).toBe(true)
    reject.click()

    expect(post).toHaveBeenCalledTimes(1)
    expect(post).toHaveBeenCalledWith({ type: 'vound:fieldAction', id: 'datum', action: 'reject' })
    expect(shown()).toBe(false)
  })

  /** a table row of two marked cells (left 196..500, right 500..806, both 100..120); returns the cells */
  function twoCellRow() {
    const post = vi.fn<(msg: EditorToHost) => void>()
    const mount = document.createElement('div')
    scroller.append(mount)
    const start = doc(para(0, 'Formular'), table(cell(para(1, 'Vergabenummer')), cell(para(2, 'dasd'))))
    view = new EditorView(mount, { state: EditorState.create({ schema, doc: start, plugins: [markersPlugin] }) })
    setMarkers(view, [{ ...marker('nummer', 1, ['fill', 'ask']), scope: 'cell' }, { ...marker('wert', 2, ['accept', 'reject']), scope: 'cell' }])
    unmount = mountFieldChip(view, post)
    const left = block(view, 'nummer')
    const right = block(view, 'wert')
    const tableElement = view.dom.querySelector('table')!
    vi.mocked(Element.prototype.getBoundingClientRect).mockImplementation(function (this: Element) {
      if (this === scroller) return rectOf({ left: 0, top: 0, right: 1200, bottom: 800 })
      if (this === view?.dom) return rectOf({ left: 100, top: 0, right: 900, bottom: 2000 })
      if (this === left) return rectOf({ left: 196, top: 100, right: 500, bottom: 120 })
      if (this === right) return rectOf({ left: 500, top: 100, right: 806, bottom: 120 })
      if (this === tableElement) return rectOf({ left: 196, top: 100, right: 806, bottom: 120 })
      if (this.classList.contains('vound-chip')) return rectOf({ left: 0, top: 0, right: 60, bottom: 30 })
      return rectOf({ left: 0, top: 0, right: 0, bottom: 0 })
    })
    return { left, right }
  }
  const tether = () => scroller.querySelector<HTMLElement>('.vound-chip-tether:not(.vound-chip-tether-elbow)')!
  const elbow = () => scroller.querySelector<HTMLElement>('.vound-chip-tether-elbow')!
  const box = (element: HTMLElement) => ({ left: element.style.left, width: element.style.width, top: element.style.top, height: element.style.height })

  it('a cell with a neighbour: the tether leaves its ring bottom, runs under the neighbour and bends up in the margin', () => {
    const { left } = twoCellRow()
    hover(left)
    expect(shown()).toBe(true)
    // from the cell's right edge along its bottom to halfway between the table (806) and the chip (910)
    expect(elbow().hidden).toBe(false)
    expect(box(elbow())).toEqual({ left: '500px', width: '358px', top: '110px', height: '10px' })
    expect(box(tether())).toEqual({ left: '858px', width: '52px', top: '110px', height: '' })
    expect(left.classList.contains('vound-slot-hot')).toBe(true)
    // on the chip its field stays ringed
    hover(chip()!.querySelector('button')!)
    expect(left.classList.contains('vound-slot-hot')).toBe(true)
  })

  it('the last cell of the row: a straight tether from the table edge, apart from its neighbour\'s', () => {
    const { left, right } = twoCellRow()
    hover(right)
    expect(elbow().hidden).toBe(true)
    expect(box(tether())).toEqual({ left: '806px', width: '104px', top: '110px', height: '' })
    expect(right.classList.contains('vound-slot-hot')).toBe(true)
    expect(left.classList.contains('vound-slot-hot')).toBe(false)
  })

  const owner = () => chip()!.querySelector<HTMLButtonElement>('button')!.dataset.id
  const caretIn = (element: Element) => {
    document.getSelection()!.collapse(element, 0)
    document.dispatchEvent(new Event('selectionchange'))
  }

  it('crossing the next cell on the way to the chip keeps the chip with the first field', () => {
    vi.useFakeTimers()
    const { left, right } = twoCellRow()
    hover(left)
    hover(right, { clientX: 600, clientY: 110 })
    vi.advanceTimersByTime(150)
    hover(chip()!.querySelector('button')!)
    vi.advanceTimersByTime(500)
    expect(shown()).toBe(true)
    expect(owner()).toBe('nummer')
    expect(left.classList.contains('vound-slot-hot')).toBe(true)
  })

  it('resting on the next cell longer than the delay switches the chip to it', () => {
    vi.useFakeTimers()
    const { left, right } = twoCellRow()
    hover(left)
    hover(right, { clientX: 600, clientY: 110 })
    // moving on inside the cell does not restart the delay
    hover(right, { clientX: 650, clientY: 112 })
    vi.advanceTimersByTime(219)
    expect(owner()).toBe('nummer')
    vi.advanceTimersByTime(1)
    expect(owner()).toBe('wert')
    expect(right.classList.contains('vound-slot-hot')).toBe(true)
    expect(left.classList.contains('vound-slot-hot')).toBe(false)
  })

  it('a field off the way to the chip takes it at once', () => {
    const { v } = chipView([marker('ort', 1, ['fill']), marker('datum', 2, ['accept'])])
    hover(block(v, 'ort'))
    hover(block(v, 'datum'), { clientX: 400, clientY: 110 })
    expect(owner()).toBe('datum')
  })

  it('the caret moving into a field shows its chip at once, out of the fields hides it', () => {
    vi.useFakeTimers()
    const { left, right } = twoCellRow()
    hover(left)
    hover(right, { clientX: 600, clientY: 110 })
    caretIn(right)
    expect(owner()).toBe('wert')
    caretIn(left)
    expect(owner()).toBe('nummer')
    caretIn(view!.dom.querySelector('p')!)
    expect(shown()).toBe(false)
  })

  it('the caret\'s chip stays when the pointer leaves for the page or the window', () => {
    vi.useFakeTimers()
    const { left } = twoCellRow()
    hover(left)
    caretIn(left)
    hover(view!.dom)
    vi.advanceTimersByTime(500)
    expect(owner()).toBe('nummer')
    document.dispatchEvent(new MouseEvent('mouseout', { relatedTarget: null }))
    vi.advanceTimersByTime(500)
    expect(owner()).toBe('nummer')
    expect(left.classList.contains('vound-slot-hot')).toBe(true)
  })

  it('leaving another field the pointer hovered brings the chip back to the caret\'s field', () => {
    vi.useFakeTimers()
    const { left, right } = twoCellRow()
    caretIn(right)
    hover(left)
    expect(owner()).toBe('nummer')
    hover(view!.dom)
    vi.advanceTimersByTime(220)
    expect(owner()).toBe('wert')
  })

  it('after a scroll the caret\'s field gets its chip back once the page rests', () => {
    vi.useFakeTimers()
    const { left } = twoCellRow()
    caretIn(left)
    scroller.dispatchEvent(new Event('scroll'))
    expect(shown()).toBe(false)
    vi.advanceTimersByTime(150)
    scroller.dispatchEvent(new Event('scroll'))
    vi.advanceTimersByTime(219)
    expect(shown()).toBe(false)
    vi.advanceTimersByTime(1)
    expect(owner()).toBe('nummer')
  })

  it('after a scroll with the caret outside the fields the chip stays hidden', () => {
    vi.useFakeTimers()
    const { left } = twoCellRow()
    hover(left)
    caretIn(view!.dom.querySelector('p')!)
    scroller.dispatchEvent(new Event('scroll'))
    vi.advanceTimersByTime(500)
    expect(shown()).toBe(false)
  })

  it('the elbow takes the missing tone and hides with the chip', () => {
    const { left } = twoCellRow()
    setMarkers(view!, [{ ...marker('nummer', 1, ['fill', 'ask']), scope: 'cell', tone: 'gap' }])
    hover(left)
    expect(elbow().classList.contains('vound-chip-tether-missing')).toBe(true)
    scroller.dispatchEvent(new Event('scroll'))
    expect(elbow().hidden).toBe(true)
  })

  it('no chip without actions', () => {
    const { v } = chipView([marker('name', 2, [])])
    hover(block(v, 'name'))
    expect(shown()).toBe(false)
  })

  it('chip hides on scroll', () => {
    const { v } = chipView([marker('name', 2, ['accept', 'reject'])])
    hover(block(v, 'name'))
    expect(shown()).toBe(true)
    scroller.dispatchEvent(new Event('scroll'))
    expect(shown()).toBe(false)
    expect(block(v, 'name').classList.contains('vound-slot-hot')).toBe(false)
  })

  it('chip hides on a window resize', () => {
    const { v } = chipView([marker('name', 2, ['accept', 'reject'])])
    hover(block(v, 'name'))
    window.dispatchEvent(new Event('resize'))
    expect(shown()).toBe(false)
  })

  it('stays while the pointer crosses the margin to the chip, hides 220ms after it left both', () => {
    vi.useFakeTimers()
    const { v } = chipView([marker('name', 2, ['accept', 'reject'])])
    hover(block(v, 'name'))
    hover(v.dom)
    vi.advanceTimersByTime(200)
    hover(chip()!.querySelector('button')!)
    vi.advanceTimersByTime(500)
    expect(shown()).toBe(true)

    hover(v.dom)
    vi.advanceTimersByTime(219)
    expect(shown()).toBe(true)
    vi.advanceTimersByTime(1)
    expect(shown()).toBe(false)
  })

  it('a multi-field chip stays inside a narrow gutter instead of being dropped', () => {
    // 100px of gutter: room for 74px of chip, enough for two buttons, the labels shrink or wrap
    gutterRight = 1000
    const { v } = chipView([marker('ort', 1, ['fill', 'ask'], 'Ort der Leistungserbringung'), marker('datum', 1, ['accept', 'reject'], 'Datum')])
    hover(block(v, 'ort'))
    expect(shown()).toBe(true)
    expect(chip()!.style.maxWidth).toBe('74px')
    expect(rowLabels()).toEqual(['Ort der Leistungserbringung', 'Datum'])
  })

  it('in a multi-column section the chip shows without its tether', () => {
    const { v } = chipView([marker('name', 2, ['accept', 'reject'])])
    v.dom.style.columnCount = '2'
    hover(block(v, 'name'))
    expect(shown()).toBe(true)
    expect(scroller.querySelector<HTMLElement>('.vound-chip-tether')!.hidden).toBe(true)
  })

  it('is not shown when the gutter is narrower than the chip, the hover outline still is', () => {
    // 75px of gutter: 49px of room, less than two 24px buttons in their frame (58px)
    gutterRight = 975
    const { v } = chipView([marker('name', 2, ['accept', 'reject'])])
    hover(block(v, 'name'))
    expect(shown()).toBe(false)
    expect(block(v, 'name').classList.contains('vound-slot-hot')).toBe(true)
  })

  it('a button whose action the markers no longer offer posts nothing', () => {
    const { v, post } = chipView([marker('name', 2, ['accept', 'reject'])])
    hover(block(v, 'name'))
    setMarkers(v, [marker('name', 2, ['reject'])])
    chip()!.querySelector<HTMLButtonElement>('[aria-label="Übernehmen"]')!.click()
    expect(post).not.toHaveBeenCalled()
  })

  it('unmounting removes the chip and its listeners', () => {
    const { v, post } = chipView([marker('name', 2, ['accept', 'reject'])])
    hover(block(v, 'name'))
    unmount!(); unmount = null
    expect(chip()).toBeNull()
    expect(markersKey.getState(v.state)!.hot).toBeNull()
    hover(block(v, 'name'))
    expect(chip()).toBeNull()
    expect(post).not.toHaveBeenCalled()
  })
})
