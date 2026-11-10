// @vitest-environment jsdom
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'

const docsCss = readFileSync(resolve(__dirname, '../../apps/docs/src/renderer/styles.css'), 'utf8')
const hostCss = readFileSync(resolve(__dirname, '../host.css'), 'utf8')

type Pseudo = '::before' | '::after'

afterEach(() => {
  document.head.innerHTML = ''
  document.body.innerHTML = ''
  document.documentElement.className = ''
})

function matchesBase(element: Element, selector: string, pseudo: Pseudo) {
  if (!selector.endsWith(pseudo)) return false
  try {
    return element.matches(selector.slice(0, -pseudo.length))
  } catch {
    return false
  }
}

/** Screen `content` of a pseudo-element by importance and source order (jsdom computes no pseudo styles). */
function pseudoContent(element: Element, pseudo: Pseudo) {
  let normal = ''
  let important = ''
  for (const sheet of Array.from(document.styleSheets)) {
    for (const rule of Array.from(sheet.cssRules)) {
      if (!(rule instanceof CSSStyleRule)) continue
      const value = rule.style.getPropertyValue('content')
      if (!value) continue
      if (!rule.selectorText.split(',').some((selector) => matchesBase(element, selector.trim(), pseudo))) continue
      if (rule.style.getPropertyPriority('content') === 'important') important = value
      else normal = value
    }
  }
  return important || normal
}

function cornerMarks(...sheets: string[]) {
  for (const css of sheets) {
    const style = document.createElement('style')
    style.textContent = css
    document.head.append(style)
  }
  const wrap = document.createElement('div')
  wrap.className = 'page-wrap'
  const gap = document.createElement('div')
  gap.className = 'page-gap'
  wrap.append(gap)
  document.body.append(wrap)
  return [pseudoContent(wrap, '::before'), pseudoContent(wrap, '::after'), pseudoContent(gap, '::after')]
}

describe('page corner marks', () => {
  it('are drawn by the docs stylesheet on its own', () => {
    expect(cornerMarks(docsCss)).not.toContain('none')
    expect(cornerMarks(docsCss)).not.toContain('')
  })

  it.each(['', 'vound-embedded', 'vound-embedded vound-review', 'vound-embedded vound-view'])(
    'are not drawn in web-full (html class "%s")',
    (mode) => {
      document.documentElement.className = mode
      expect(cornerMarks(docsCss, hostCss)).toEqual(['none', 'none', 'none'])
    },
  )
})
