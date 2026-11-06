/**
 * The last page's footnotes are not in a page gap: they render after the document. They must end on the
 * page's bottom text boundary (--page-mb), above the footer, not at a fixed pull-up that ignores the footer:
 * a footer with a top rule then crossed the last footnote line. Endnotes keep their own placement.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')
const ruleOf = (selector: string) =>
  css.match(new RegExp(`(?:^|\\n)${selector.replace(/[.()+]/g, '\\$&')} \\{([^}]*)\\}`))?.[1] ??
  null

describe('last-page footnotes', () => {
  it('sit on the bottom text boundary of the page', () => {
    const rule = ruleOf('.page-notes:not(.page-endnotes)')
    expect(rule).not.toBeNull()
    expect(rule).toContain('position: absolute')
    expect(rule).toContain('bottom: var(--page-mb, 96px)')
    expect(rule).toContain('padding-bottom: 0')
  })

  it('leave the endnote area as it was', () => {
    expect(ruleOf('.page-notes + .page-endnotes')).toContain('margin-top: 0')
    expect(ruleOf('.page-endnotes-anchored')).toContain('position: absolute')
  })
})
