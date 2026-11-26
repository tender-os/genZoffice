import { describe, expect, it } from 'vitest'
import { countByPage, pageListLabels, pageTitle } from '../pageCounts'

const pages = [
  { top: 0, height: 1000 },
  { top: 1028, height: 872 },
  { top: 1900, height: 600 },
]

describe('countByPage', () => {
  it('counts each field on the page it starts on, by what the reviewer still has to do', () => {
    const counts = countByPage(pages, [
      { tone: 'proposed', top: 100 },
      { tone: 'accepted', top: 200 },
      { tone: 'gap', top: 1900 },
      { tone: 'conditional', top: 2000 },
      { tone: 'refused', top: 2100 },
      { tone: 'rejected', top: 2200 },
      { tone: 'otherParty', top: 2300 },
    ])
    expect(counts).toEqual([
      { fields: 2, missing: 0, toReview: 1 },
      { fields: 0, missing: 0, toReview: 0 },
      { fields: 5, missing: 3, toReview: 0 },
    ])
  })
})

describe('pageTitle', () => {
  const de = pageListLabels('de')
  it('tells a page without fields from one whose fields are all decided', () => {
    expect(pageTitle({ fields: 0, missing: 0, toReview: 0 }, de)).toBe('Keine Eingaben nötig')
    expect(pageTitle({ fields: 4, missing: 0, toReview: 0 }, de)).toBe('Vollständig ausgefüllt')
  })
  it('names only what is open', () => {
    expect(pageTitle({ fields: 4, missing: 1, toReview: 0 }, de)).toBe('1 fehlt')
    expect(pageTitle({ fields: 4, missing: 2, toReview: 3 }, de)).toBe('2 fehlen · 3 zu prüfen')
  })
})
