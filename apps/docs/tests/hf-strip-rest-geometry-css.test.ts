/**
 * A strip's reserved height is measured without the editor-chrome floors (hf-dom.ts hfProbeHost), and the
 * footer strip is pinned by its bottom at the footer distance. Drawn with the 1.3em paragraph floor, a footer
 * rose above the body's bottom boundary by the difference: its top rule crossed the corner marks and the
 * footnote area ending there. The page-gap footnote area must end on the boundary too, not 5px past it, and the
 * crop marks in a page gap sit on that boundary, below the footnotes and padding the gap carries above it.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')
const ruleOf = (selector: string) =>
  css.match(new RegExp(`(?:^|\\n)${selector.replace(/[.()+:]/g, '\\$&')} \\{([^}]*)\\}`))?.[1] ??
  null

describe('header/footer strips at rest', () => {
  it('draw their paragraphs at the measured geometry, without the editing floor', () => {
    expect(ruleOf('.page-hf:not(.page-hf-editing) .page-hf-para')).toContain('min-height: 0')
    // the floor still keeps the lines of a strip being edited clickable
    expect(ruleOf('.page-hf-para')).toContain('min-height: 1.3em')
  })
})

describe('page-gap footnote area', () => {
  it('starts on the body boundary, so it ends where its reserved height ends', () => {
    expect(ruleOf('.page-gap-notes')).toMatch(/\btop: 0;/)
  })
})

describe('page-gap crop marks', () => {
  it('sit on the previous page text boundary, not where its content ends', () => {
    const marks = ruleOf('.page-gap:not(.page-gap-inline)::after') ?? ''
    expect(marks).toContain(
      '--cm-y: calc(var(--gap-mb, 96px) - var(--gap-page-mb, var(--gap-mb, 96px)))',
    )
    expect(marks).toContain('var(--gap-ml, 96px) var(--cm-y)')
  })
})
