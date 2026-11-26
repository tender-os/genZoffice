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
      { tone: 'proposed', y: 100 },
      { tone: 'accepted', y: 200 },
      { tone: 'gap', y: 1900 },
      { tone: 'conditional', y: 2000 },
      { tone: 'refused', y: 2100 },
      { tone: 'rejected', y: 2200 },
      { tone: 'otherParty', y: 2300 },
    ])
    expect(counts).toEqual([
      { fields: 2, missing: 0, toReview: 1 },
      { fields: 0, missing: 0, toReview: 0 },
      { fields: 5, missing: 3, toReview: 0 },
    ])
  })
})

describe('the page a field is counted on', () => {
  it('is the page below for a line drawn in the band between two papers', () => {
    // page 1's paper ends at 1000, page 2's starts at 1028
    const counts = countByPage(pages, [{ tone: 'gap', y: 1010 }])
    expect(counts.map((page) => page.missing)).toEqual([0, 1, 0])
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
