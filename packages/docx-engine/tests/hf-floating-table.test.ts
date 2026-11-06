import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/** header/footer tables floated against the page or the margins (w:tblpPr) */

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const FTR_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'

const cell = (pct: number, text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="${pct}" w:type="pct"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`

/** a footer table pinned to the page bottom across the text column, then the part's empty paragraph */
const floatingTable = (tblpPr: string, tblW = '<w:tblW w:w="5000" w:type="pct"/>') =>
  `<w:tbl><w:tblPr><w:tblpPr w:leftFromText="187" w:rightFromText="187" ${tblpPr}/>${tblW}<w:tblLayout w:type="fixed"/></w:tblPr>` +
  '<w:tblGrid><w:gridCol w:w="8618"/><w:gridCol w:w="2154"/></w:tblGrid>' +
  `<w:tr><w:trPr><w:trHeight w:val="727"/></w:trPr>${cell(4000, 'Changed 10/2026')}${cell(1000, '1')}</w:tr></w:tbl><w:p/>`

async function footerParas(partXml: string) {
  const bytes = await buildDocx({
    bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    extraRels:
      '<Relationship Id="rId61" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
    extraParts: [
      { path: 'word/footer1.xml', xml: `<w:ftr ${W}>${partXml}</w:ftr>`, contentType: FTR_CT },
    ],
    sectPrExtra: '<w:footerReference w:type="default" r:id="rId61"/>',
  })
  const paras = (await parseDocx(bytes)).footerParas!
  return { row: paras.find((para) => para.cells)!, flow: paras.filter((para) => !para.cells) }
}

describe('floating header/footer tables', () => {
  it('draw a page-anchored table at its page position, off the strip flow', async () => {
    const { row, flow } = await footerParas(
      floatingTable(
        'w:vertAnchor="page" w:horzAnchor="margin" w:tblpXSpec="center" w:tblpYSpec="bottom"',
      ),
    )
    expect(row.cells).toHaveLength(2)
    expect(row.boxAnchored).toBe(true)
    expect(row.box).toMatchObject({
      posVRel: 'page',
      posV: 'bottom',
      posHRel: 'margin',
      posH: 'center',
      insets: [0, 0, 0, 0],
      widthPct: 100,
    })
    // the empty paragraph after the table stays the strip's only flow line
    expect(flow).toHaveLength(1)
    expect(flow[0].boxAnchored).toBeUndefined()
  })

  it('take absolute offsets against the margins, and a fixed width', async () => {
    const { row } = await footerParas(
      floatingTable(
        'w:vertAnchor="margin" w:horzAnchor="page" w:tblpX="1500" w:tblpY="300"',
        '<w:tblW w:w="6000" w:type="dxa"/>',
      ),
    )
    expect(row.box).toMatchObject({
      posVRel: 'margin',
      posHRel: 'page',
      posXPx: 100,
      posYPx: 20,
      widthPx: 400,
    })
    expect(row.box?.posV).toBeUndefined()
  })

  it('keep a table floated against the text in the strip flow', async () => {
    const { row } = await footerParas(floatingTable('w:vertAnchor="text" w:tblpY="120"'))
    expect(row.cells).toHaveLength(2)
    expect(row.box).toBeUndefined()
    expect(row.boxAnchored).toBeUndefined()
  })
})
