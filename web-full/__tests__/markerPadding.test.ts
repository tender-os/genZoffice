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

describe('review field markers', () => {
  it.each(['p', 'td'] as const)(
    'leave the box of a marked %s as it is, in every tone and state',
    (tag) => {
      const plain = boxOf(block(tag))
      for (const tone of TONES) {
        for (const state of ['', 'vound-slot-current', 'vound-slot-hot']) {
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
    expect(variable(paragraph, '--vs-plate')).toBe(
      '0 0 0 1px var(--vs-tint), -2px 0 0 1px var(--vs-tint), 2px 0 0 1px var(--vs-tint), -4px 0 0 1px var(--vs-bar)',
    )
    expect(getComputedStyle(paragraph).getPropertyValue('background')).toBe('var(--vs-tint)')
  })

  it('give a conditional paragraph a single lighter bar in the place of the others', () => {
    const paragraph = block('p', 'vound-slot vound-slot-conditional')
    expect(variable(paragraph, '--vs-l')).toBe('5px')
    expect(variable(paragraph, '--vs-plate')).toBe(
      '0 0 0 1px var(--vs-tint), -2px 0 0 1px var(--vs-tint), 2px 0 0 1px var(--vs-tint), -4px 0 0 1px var(--vound-conditional-bar)',
    )
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

  it('keep the dashed bar of a conditional cell', () => {
    const cell = block('td', 'vound-slot vound-slot-conditional')
    expect(variable(cell, '--vs-bar-image')).toContain('repeating-linear-gradient')
    expect(variable(cell, '--vs-plate')).toBe('0 0 0 0 transparent')
  })
})
