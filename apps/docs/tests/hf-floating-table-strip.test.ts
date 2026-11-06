// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import type { HfParagraph, HfTextBox } from '@genoffice/docx-engine'
import {
  hfReservedHeightPx,
  hfTextBoxStyle,
  makeGapHfEl,
  type HfStripGeom,
} from '../src/renderer/editor/hf-dom'
import { estimateHfHeight } from '../src/renderer/line-metrics'

/** A4 with 0.39in margins (567 twips) and the footer 709 twips above the page bottom (px) */
const geom: HfStripGeom = {
  pageW: 793.7,
  pageH: 1122.5,
  marginLeft: 37.8,
  marginRight: 37.8,
  marginTop: 37.8,
  marginBottom: 37.8,
  headerStripTop: 47.27,
  footerDist: 47.27,
}
const contentW = geom.pageW - geom.marginLeft - geom.marginRight

/** a footer table floated to the page bottom across the text column (w:tblpPr vertAnchor page, tblpYSpec bottom) */
const tableBox: HfTextBox = {
  id: 7,
  insets: [0, 0, 0, 0],
  posVRel: 'page',
  posV: 'bottom',
  posHRel: 'margin',
  posH: 'center',
  widthPct: 100,
}
const emptyPara: HfParagraph = { runs: [], lineRule: 'auto', lineRawTwips: 240, lineSpacing: 1 }
const floatingRow: HfParagraph = {
  runs: [],
  boxAnchored: true,
  box: tableBox,
  row: { heightTwips: 727, heightRule: 'atLeast' },
  cells: [
    { paras: [[{ text: 'Changed 10/2026', sizeHalfPoints: 12 }]], widthPct: 80 },
    { paras: [[{ text: '1', sizeHalfPoints: 28 }]], widthPct: 20 },
  ],
}

describe('a footer table floated against the page', () => {
  it('spans the text column with its bottom on the page bottom', () => {
    const style = hfTextBoxStyle(tableBox, 'footer', geom)!
    expect(style.width).toBe('100%')
    // the strip ends footerDist above the page bottom; the box ends footerDist lower
    expect(parseFloat(style.bottom)).toBeCloseTo(-geom.footerDist, 6)
    expect(style.transform).toBe('translate(-50%, 0%)')
  })

  it('reserves only the strip flow: the empty paragraph after the table', () => {
    const value = { text: '', paras: [emptyPara, floatingRow] }
    expect(hfReservedHeightPx('footer', value, contentW)).toBe(
      estimateHfHeight({ text: '', paras: [emptyPara] }, contentW),
    )
  })

  it('draws the rows in the positioned box', () => {
    const el = makeGapHfEl({
      kind: 'footer',
      value: { text: '', paras: [emptyPara, floatingRow] },
      pageNo: 1,
      pageTotal: 1,
      geom,
    })
    const box = el.querySelector<HTMLElement>('.page-hf-textbox')
    expect(box?.querySelector('.page-hf-row')).not.toBeNull()
    expect(box?.style.width).toBe('100%')
  })
})
