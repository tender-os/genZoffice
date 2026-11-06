// @vitest-environment jsdom
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { ringsIn } from '../fieldRing'

const hostCss = readFileSync(resolve(__dirname, '../host.css'), 'utf8')
const origin = { left: 0, top: 0 }

beforeEach(() => {
  const style = document.createElement('style')
  style.textContent = hostCss
  document.head.append(style)
  document.documentElement.classList.add('vound-review')
})
afterEach(() => {
  document.head.innerHTML = ''
  document.body.innerHTML = ''
  document.documentElement.classList.remove('vound-review')
})

/** a block at x 100..500, rows 20px apart starting at `top`, 16px tall */
function place(element: Element, top: number, left = 100, width = 400) {
  element.getBoundingClientRect = () => new DOMRect(left, top, width, 16)
  return element
}

function marked(tag: string, ids: string, extra = '') {
  const element = document.createElement(tag)
  element.className = `vound-slot vound-slot-gap ${extra}`.trim()
  element.setAttribute('data-vound-slot', ids)
  return element
}

/** paragraphs (or list items) from specs: [ids | null, classes] */
function blocks(tag: string, specs: Array<[string | null, string?]>) {
  const root = document.createElement('div')
  specs.forEach(([ids, extra], i) => root.append(place(ids === null ? document.createElement(tag) : marked(tag, ids, extra), 20 * i)))
  document.body.append(root)
  return root
}

const boxes = (root: Element) => ringsIn(root, origin).map(({ kind, left, top, width, height }) => ({ kind, left, top, width, height }))

describe('field rings', () => {
  it('draws one closed ring around a field that spans two paragraphs, past their tint', () => {
    const root = blocks('p', [['a', 'vound-slot-current'], ['a', 'vound-slot-current'], [null]])
    // tint reaches 5px left, 3px right, 1px above and below each paragraph
    expect(boxes(root)).toEqual([{ kind: 'current', left: 95, top: -1, width: 408, height: 38 }])
  })

  it('draws one ring around three list items, and a closed one around a single block', () => {
    const root = blocks('div', [['a', 'vound-slot-current'], ['a', 'vound-slot-current'], ['a', 'vound-slot-current'], [null], ['b', 'vound-slot-current']])
    expect(boxes(root)).toEqual([
      { kind: 'current', left: 95, top: -1, width: 408, height: 58 },
      { kind: 'current', left: 95, top: 79, width: 408, height: 18 },
    ])
  })

  it('keeps adjacent blocks of different fields, and runs of one field split by other text, apart', () => {
    const root = blocks('p', [['a', 'vound-slot-current'], ['b'], ['a', 'vound-slot-current'], [null], ['a', 'vound-slot-current']])
    expect(boxes(root).map((ring) => ring.top)).toEqual([-1, 39, 79])
  })

  it('rings the whole run of the hovered block, and no hover ring on a selected one', () => {
    const root = blocks('p', [['a'], ['a', 'vound-slot-hot'], ['a'], ['b', 'vound-slot-current']])
    expect(boxes(root)).toEqual([
      { kind: 'current', left: 95, top: 59, width: 408, height: 18 },
      { kind: 'hot', left: 95, top: -1, width: 408, height: 58 },
    ])
    const selected = blocks('p', [['c', 'vound-slot-current vound-slot-hot'], ['c', 'vound-slot-current']])
    expect(boxes(selected).map((ring) => ring.kind)).toEqual(['current'])
  })

  it('joins the cells of one column in consecutive rows, and only those', () => {
    const table = document.createElement('table')
    const rows = [0, 1, 2].map((i) => {
      const row = table.insertRow()
      row.append(place(document.createElement('td'), 20 * i, 100, 200))
      row.append(place(marked('td', i < 2 ? 'a' : 'b', 'vound-slot-current'), 20 * i, 300, 200))
      return row
    })
    rows[0].cells[0].replaceWith(place(marked('td', 'a', 'vound-slot-current'), 0, 100, 200))
    document.body.append(table)
    // a cell's tint stays inside it: the ring goes round the cells
    expect(boxes(table)).toEqual([
      { kind: 'current', left: 100, top: 0, width: 200, height: 16 },
      { kind: 'current', left: 300, top: 0, width: 200, height: 36 },
      { kind: 'current', left: 300, top: 40, width: 200, height: 16 },
    ])
  })

  it('leaves the ring off the block itself, where the next marked block would paint over it', () => {
    const [plain, selected, hovered] = [marked('p', 'a'), marked('p', 'a', 'vound-slot-current'), marked('p', 'a', 'vound-slot-hot')]
    document.body.append(plain, selected, hovered)
    expect(getComputedStyle(selected).boxShadow).toBe(getComputedStyle(plain).boxShadow)
    expect(getComputedStyle(hovered).boxShadow).toBe(getComputedStyle(plain).boxShadow)
  })
})
