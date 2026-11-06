import { describe, expect, it } from 'vitest'
import type { SectionInfo } from '@genoffice/docx-engine'
import {
  columnLayoutSpecs,
  computeSectionedSlicesF2,
  sectionColGeom,
  stackedPullUpPx,
  type BlockBox,
  type ColumnBlockPlacement,
  type PageSlice,
  type SectionGeom,
  type TableRowBox,
} from '../src/renderer/pagination'

const sec = (columns: number, startType: SectionInfo['startType']): SectionInfo => ({
  settings: {
    pageWidth: 11906,
    pageHeight: 16838,
    orientation: 'portrait',
    marginTop: 1134,
    marginRight: 720,
    marginBottom: 1134,
    marginLeft: 720,
    pageBorder: false,
    columns,
    colSpace: 402,
  },
  startType,
  firstBlockIndex: 0,
  lastBlockIndex: 0,
  sectPrXml: '',
  titlePg: false,
  headerRefs: {},
  footerRefs: {},
})

const CONTENT_H = 900
const TITLE_H = 100
const sections = [sec(1, 'nextPage'), sec(2, 'continuous')]
const geoms: SectionGeom[] = [
  { contentHeight: CONTENT_H, forceBreak: false, cols: 1 },
  { contentHeight: CONTENT_H, forceBreak: false, cols: 2, startType: 'continuous' },
]

/** a one-column title, then a continuous two-column section holding one table */
function layout(rows: TableRowBox[]) {
  const height = rows.reduce((s, r) => s + r.height, 0)
  const blocks: BlockBox[] = [
    { top: 0, height: TITLE_H, section: 0, el: document.createElement('p') },
    {
      top: TITLE_H,
      height,
      section: 1,
      el: document.createElement('table'),
      tableRows: rows,
    },
  ]
  const slices = computeSectionedSlicesF2(blocks, geoms, TITLE_H + height)
  const spec = columnLayoutSpecs(blocks, slices, sections).find((s) => s.el === blocks[1].el)
  return { blocks, slices, spec: spec as ColumnBlockPlacement }
}

type Column = {
  page: number
  col: number
  start: number
  end: number
  top: number
  height: number
  header: number
}

function engineColumns(slices: PageSlice[]): Column[] {
  return slices.flatMap((s, page) =>
    (s.regions ?? []).flatMap((r) =>
      r.columns.map((c, col) => ({
        page,
        col: r.columns.length > 1 ? col : -1,
        start: c.start,
        end: c.end,
        top: r.top,
        height: r.height,
        header: c.repeatHeader?.height ?? 0,
      })),
    ),
  )
}

/**
 * Where the canvas paints each row, from the DOM model the editor uses: the
 * stacked single flow plus the page gaps' repeated header rows, shifted by the
 * table's column translate and its row fragment's, against the physical page.
 */
function canvasPlacement(rows: TableRowBox[], slices: PageSlice[], spec: ColumnBlockPlacement) {
  const xs = sectionColGeom(sections[1])
  const colX = [0, xs.widths[0] + xs.gaps[0]]
  const tablePage = slices.findIndex((s) => TITLE_H >= s.start && TITLE_H < s.end)
  let top = TITLE_H
  return rows.map((row, ri) => {
    const frag = spec.rowFragments?.find((f) => ri >= f.startRow && ri < f.endRow)
    const dx = spec.dx + (frag?.dx ?? 0)
    const dy = spec.dy + (frag?.dy ?? 0)
    const col = Math.abs(dx - colX[0]) < Math.abs(dx - colX[1]) ? 0 : 1
    let page = tablePage
    let domY = top - slices[tablePage].start
    while (page + 1 < slices.length && domY + dy >= slices[page].physHeight! - 0.5) {
      domY -= slices[page].physHeight!
      page++
    }
    const gapHeaders = slices
      .slice(tablePage + 1, page + 1)
      .reduce((s, sl) => s + (sl.repeatHeader?.height ?? 0), 0)
    const y = domY + gapHeaders + dy
    top += row.height
    return { ri, page, col, y, bottom: y + row.height }
  })
}

function previewPlacement(rows: TableRowBox[], slices: PageSlice[]) {
  const cols = engineColumns(slices).filter((c) => c.col >= 0)
  let top = TITLE_H
  return rows.map((row, ri) => {
    const c = cols.find((k) => top >= k.start - 0.5 && top < k.end - 0.5)!
    const y = c.top + c.header + (top - c.start)
    top += row.height
    return { ri, page: c.page, col: c.col, y, bottom: y + row.height, limit: c.top + c.height }
  })
}

const plainRows = (n: number): TableRowBox[] => Array.from({ length: n }, () => ({ height: 50 }))

describe('column-split table (row-level placement in mixed-column canvas)', () => {
  it('cuts a one-column-wide table at row boundaries into the next column and page', () => {
    const rows = plainRows(56)
    const { slices, spec } = layout(rows)
    expect(slices).toHaveLength(2)
    const preview = previewPlacement(rows, slices)
    const canvas = canvasPlacement(rows, slices, spec)
    // engine columns: page 1 → 16 + 16 rows, page 2 → 18 + 6 rows
    expect(preview.map((p) => `${p.page}/${p.col}`)).toEqual([
      ...Array(16).fill('0/0'),
      ...Array(16).fill('0/1'),
      ...Array(18).fill('1/0'),
      ...Array(6).fill('1/1'),
    ])
    expect(spec.rowFragments?.map((f) => [f.startRow, f.endRow])).toEqual([
      [16, 32],
      [32, 50],
      [50, 56],
    ])
    for (const c of canvas) {
      const p = preview[c.ri]
      expect({ ri: c.ri, page: c.page, col: c.col, y: Math.round(c.y) }).toEqual({
        ri: p.ri,
        page: p.page,
        col: p.col,
        y: Math.round(p.y),
      })
      expect(p.bottom).toBeLessThanOrEqual(p.limit + 0.5)
    }
  })

  it('repeats header rows at every column top and keeps a cantSplit row whole', () => {
    const rows: TableRowBox[] = plainRows(56).map((r, i) =>
      i === 0 ? { ...r, isHeader: true } : i === 14 ? { height: 150, cantSplit: true } : r,
    )
    const { slices, spec } = layout(rows)
    const preview = previewPlacement(rows, slices)
    const canvas = canvasPlacement(rows, slices, spec)
    // the cantSplit row does not fit page 1's first column: it opens the second whole
    expect(preview[14]).toMatchObject({ page: 0, col: 1 })
    expect(preview[14].y).toBe(TITLE_H + 50)
    // columns after the first repeat the header: painted by the canvas, except where
    // the page gap already clones it into the flow
    const painted = spec.rowFragments?.map((f) => f.headerPx)
    expect(painted).toEqual([50, 0, 50])
    for (const c of canvas) {
      const p = preview[c.ri]
      expect([c.ri, c.page, c.col, Math.round(c.y)]).toEqual([p.ri, p.page, p.col, Math.round(p.y)])
      expect(p.bottom).toBeLessThanOrEqual(p.limit + 0.5)
    }
  })

  it('keeps a vMerge continuation row with the row it continues', () => {
    const rows: TableRowBox[] = plainRows(40)
    rows[16] = { height: 50, vMergeContinue: true }
    const { slices, spec } = layout(rows)
    const frag = spec.rowFragments?.find((f) => f.startRow <= 16 && f.endRow > 16)
    const head = spec.rowFragments?.find((f) => f.startRow <= 15 && f.endRow > 15)
    expect(frag).toEqual(head)
    expect(slices[0].regions?.[1].columns[1].start).toBeGreaterThan(TITLE_H + 16 * 50)
  })

  it('pulls the flow after a table that turns pages by the vacated column space', () => {
    const { slices } = layout(plainRows(56))
    expect(stackedPullUpPx(slices[0])).toBeCloseTo(
      slices[0].end - slices[0].start - slices[0].physHeight!,
      3,
    )
    expect(stackedPullUpPx({ start: 0, end: 900, section: 0 })).toBe(0)
  })

  it('leaves a table that fits its column alone', () => {
    const { spec } = layout(plainRows(10))
    expect(spec.rowFragments).toBeUndefined()
  })
})
