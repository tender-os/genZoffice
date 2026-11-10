// @vitest-environment jsdom
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

const hostCss = readFileSync(resolve(__dirname, '../host.css'), 'utf8')
const TONES = ['proposed', 'gap', 'conditional', 'refused', 'accepted', 'rejected', 'otherParty']
const BOX_PROPERTIES = [
  'display',
  'position',
  'width',
  'height',
  'line-height',
  'padding-top',
  'padding-right',
  'padding-bottom',
  'padding-left',
  'margin-top',
  'margin-right',
  'margin-bottom',
  'margin-left',
  'border-top-width',
  'border-right-width',
  'border-bottom-width',
  'border-left-width',
  'letter-spacing',
  'font-size',
  'transform',
]

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

function block(tag: 'p' | 'td', className = '') {
  const table = document.createElement('table')
  const node = document.createElement(tag)
  node.className = className
  node.textContent = 'value'
  if (tag === 'td') table.insertRow().append(node)
  document.body.append(tag === 'td' ? table : node)
  return node
}

const boxOf = (node: HTMLElement) => {
  const style = getComputedStyle(node)
  return BOX_PROPERTIES.map((name) => `${name}: ${style.getPropertyValue(name)}`)
}

const variable = (node: HTMLElement, name: string) =>
  getComputedStyle(node).getPropertyValue(name).trim()

/** the plate's shadow layers with the frame and the bar's end put in */
const plate = (node: HTMLElement) =>
  variable(node, '--vs-plate')
    .replace('var(--vs-frame)', variable(node, '--vs-frame'))
    .replace('var(--vs-bar-end)', variable(node, '--vs-bar-end'))
/** a paragraph's tint: an image layer, so it can stop short of an edge */
const TINT_LAYER = 'linear-gradient(var(--vs-tint), var(--vs-tint))'

/** a bar layer `x 0 0 spread colour`: its left edge and how far it is shortened at each end */
const dashLayer = (shadow: string) => {
  const [x, , , spread, colour] = shadow.split(' ')
  return { left: parseFloat(x) - parseFloat(spread), shrink: 1 - parseFloat(spread), colour }
}

describe('review field markers', () => {
  it.each(['p', 'td'] as const)(
    'leave the box of a marked %s as it is, in every tone and state',
    (tag) => {
      const plain = boxOf(block(tag))
      for (const tone of TONES) {
        for (const state of ['', 'vound-slot-current', 'vound-slot-hot', 'vound-slot-other-above vound-slot-other-below', 'vound-slot-continued']) {
          const marked = block(tag, `vound-slot vound-slot-${tone} vound-slot-field ${state}`)
          expect(boxOf(marked)).toEqual(plain)
        }
      }
    },
  )

  it('draw a paragraph tint 3px beyond its text inline and 1px above and below, with the bar outside it', () => {
    const paragraph = block('p', 'vound-slot vound-slot-gap')
    expect([
      variable(paragraph, '--vs-l'),
      variable(paragraph, '--vs-r'),
      variable(paragraph, '--vs-y'),
    ]).toEqual(['5px', '3px', '1px'])
    expect(plate(paragraph)).toBe(
      '0 0 0 1px var(--vs-tint), -2px 0 0 1px var(--vs-tint), 2px 0 0 1px var(--vs-tint), -4px 0 0 1px var(--vs-bar)',
    )
    expect(getComputedStyle(paragraph).getPropertyValue('background')).toBe(TINT_LAYER)
  })

  it.each([
    ['vound-slot-other-below', '0 -1px 0 0', '-3px -1px 0 0', '3px -1px 0 0', '-5px -1px 0 0', 'center top'],
    ['vound-slot-other-above', '0 1px 0 0', '-3px 1px 0 0', '3px 1px 0 0', '-5px 1px 0 0', 'center bottom'],
  ])('stop a paragraph tint and bar 1px short of the edge another field touches (%s)', (side, ...expected) => {
    const [frame, left, right, bar, position] = expected
    const paragraph = block('p', `vound-slot vound-slot-gap ${side}`)
    expect(plate(paragraph)).toBe(
      `${frame} var(--vs-tint), ${left} var(--vs-tint), ${right} var(--vs-tint), ${bar} var(--vs-bar)`,
    )
    // jsdom drops the calc() of background-size; the browser renders check the 1px
    const style = getComputedStyle(paragraph)
    expect([style.getPropertyValue('background-position'), style.getPropertyValue('background-repeat')]).toEqual([
      position,
      'no-repeat',
    ])
  })

  it('stop the tint and bar 1px short of both edges between two other fields', () => {
    const paragraph = block('p', 'vound-slot vound-slot-conditional vound-slot-other-above vound-slot-other-below')
    const layers = plate(paragraph).split(/,\s*/)
    expect(layers.slice(0, 3)).toEqual(['0 0 0 -1px var(--vs-tint)', '-4px 0 0 -1px var(--vs-tint)', '4px 0 0 -1px var(--vs-tint)'])
    expect(layers.at(-1)).toBe('-6px 0 0 -1px var(--vs-bar)')
    const style = getComputedStyle(paragraph)
    expect([style.getPropertyValue('background-position'), style.getPropertyValue('background-repeat')]).toEqual([
      'center center',
      'no-repeat',
    ])
  })

  it('give a conditional paragraph a dashed bar in the place of the others, over a hatched tint', () => {
    const paragraph = block('p', 'vound-slot vound-slot-conditional')
    expect(variable(paragraph, '--vs-l')).toBe('5px')
    const layers = plate(paragraph).split(/,\s*/)
    expect(layers.slice(0, 3)).toEqual([
      '0 0 0 1px var(--vs-tint)',
      '-2px 0 0 1px var(--vs-tint)',
      '2px 0 0 1px var(--vs-tint)',
    ])
    const bar = layers.slice(3).reverse().map(dashLayer)
    expect(bar[0]).toEqual({ left: -5, shrink: 0, colour: 'var(--vs-bar)' })
    for (const layer of bar) expect(layer.left).toBe(-5)
    const steps = bar.slice(1).map((layer, i) => layer.shrink - bar[i].shrink)
    expect(steps).toEqual(steps.map((_, i) => (i % 2 === 0 ? 4 : 3)))
    expect(bar.map((layer) => layer.colour)).toEqual(
      bar.map((_, i) => (i % 2 === 0 ? 'var(--vs-bar)' : 'var(--vs-tint)')),
    )
    expect(bar.at(-1)!.shrink).toBeGreaterThanOrEqual(60)
    expect(getComputedStyle(paragraph).getPropertyValue('background')).toBe(`var(--vs-hatch), ${TINT_LAYER}`)
    expect(variable(paragraph, '--vs-hatch')).toContain(
      'repeating-linear-gradient(-45deg, transparent 0, var(--vound-conditional-hatch)',
    )
  })

  it("give the reviewer's own text a quiet tint and a solid bar, never the conditional dashes or hatch", () => {
    const paragraph = block('p', 'vound-slot vound-slot-conditional vound-slot-field vound-slot-own-text')
    expect(variable(paragraph, '--vs-tint')).toBe('var(--vound-own-text-tint)')
    expect(variable(paragraph, '--vs-bar')).toBe('var(--vound-own-text-bar)')
    expect(variable(paragraph, '--vs-plate')).toBe('var(--vs-frame), var(--vs-bar-end)')
    const cell = block('td', 'vound-slot vound-slot-conditional vound-slot-field vound-slot-own-text')
    expect(getComputedStyle(cell).getPropertyValue('background')).not.toContain('--vs-hatch')
  })

  it('keep a table cell marker inside the cell, where the cell padding already keeps text off the bar', () => {
    const cell = block('td', 'vound-slot vound-slot-gap')
    expect([variable(cell, '--vs-l'), variable(cell, '--vs-r'), variable(cell, '--vs-y')]).toEqual([
      '0px',
      '0px',
      '0px',
    ])
    expect(variable(cell, '--vs-plate')).toBe('0 0 0 0 transparent')
    expect(getComputedStyle(cell).getPropertyValue('background')).toContain('var(--vs-bar-image)')
  })

  it('give a conditional cell the dashed bar over the same hatched tint', () => {
    const cell = block('td', 'vound-slot vound-slot-conditional')
    expect(variable(cell, '--vs-bar-image')).toContain('repeating-linear-gradient')
    expect(variable(cell, '--vs-plate')).toBe('0 0 0 0 transparent')
    expect(getComputedStyle(cell).getPropertyValue('background')).toBe(
      'var(--vs-bar-image) left top / 2px 100% no-repeat, var(--vs-hatch), var(--vs-tint)',
    )
    expect(variable(cell, '--vs-hatch')).toBe(
      variable(block('p', 'vound-slot vound-slot-conditional'), '--vs-hatch'),
    )
  })
})
