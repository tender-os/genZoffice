/**
 * Source paragraph index: every block and table cell paragraph built from a
 * body paragraph carries the position of its w:p among all elements named `p`
 * in the original document.xml (mc:Fallback, comments and CDATA excluded).
 * A host names paragraphs by this index; nothing of it reaches a saved file.
 */
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import {
  paragraphStartOffsets,
  parseDocx,
  saveDocx,
  type Block,
  type TableCell,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const MC_NS = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'

const P = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
const TC = (inner: string, tcPr = '') => `<w:tc>${tcPr}${inner}</w:tc>`
const TABLE = (rows: string[], cols = 2) =>
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  `<w:tblGrid>${'<w:gridCol w:w="3000"/>'.repeat(cols)}</w:tblGrid>` +
  rows.map((row) => `<w:tr>${row}</w:tr>`).join('') +
  '</w:tbl>'

/** a paragraph holding a whole legacy text form field: a protected passthrough block */
const FIELD_P =
  '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>entry</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'

const TEXTBOX_ANCHOR =
  '<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>9525</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="1000000" cy="152400"/><wp:wrapNone/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="152400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  '<wps:txbx><w:txbxContent><w:p><w:r><w:t>box text</w:t></w:r></w:p></w:txbxContent></wps:txbx>' +
  '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>'

const TRIANGLE_ANCHOR =
  '<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="589915" cy="533400"/><wp:wrapNone/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="589915" cy="533400"/></a:xfrm>' +
  '<a:prstGeom prst="triangle"><a:avLst/></a:prstGeom></wps:spPr>' +
  '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>'

const visible = (blocks: Block[]) => blocks.filter((b) => !b.hidden)
const blockWithText = (blocks: Block[], text: string) =>
  blocks.find((b) => b.runs?.some((r) => r.text.includes(text)))
const cellIndices = (cell: TableCell) => cell.richParas!.map((p) => p.sourceParaIndex)

const TABLE_DOC_BODY =
  '<w:p/>' +
  TABLE([
    TC(P('r0c0')) + TC(P('r0c1'), '<w:tcPr><w:vMerge w:val="restart"/></w:tcPr>'),
    TC(P('r1c0a') + P('r1c0b')) + TC('<w:p/>', '<w:tcPr><w:vMerge/></w:tcPr>'),
  ]) +
  FIELD_P +
  '<w:p/>'

describe('source paragraph index', () => {
  it('numbers top-level paragraphs in document order', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: P('a') + P('b') + P('c') }))
    expect(visible(doc.blocks).map((b) => b.sourceParaIndex)).toEqual([0, 1, 2])
  })

  it('numbers a field paragraph that becomes a passthrough block', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: P('before') + FIELD_P + P('after') }))
    const blocks = visible(doc.blocks)
    expect(blocks.map((b) => b.type)).toEqual(['paragraph', 'passthrough', 'paragraph'])
    expect(blocks.map((b) => b.sourceParaIndex)).toEqual([0, 1, 2])
  })

  it('numbers cell paragraphs through the table, including a merged-away cell and a two-paragraph cell', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: TABLE_DOC_BODY }))
    const blocks = visible(doc.blocks)
    expect(blocks.map((b) => b.type)).toEqual(['paragraph', 'table', 'passthrough', 'paragraph'])
    expect(blocks[0].sourceParaIndex).toBe(0)
    const rows = blocks[1].table!.rows
    expect(blocks[1].sourceParaIndex).toBeUndefined()
    expect(cellIndices(rows[0][0])).toEqual([1])
    expect(cellIndices(rows[0][1])).toEqual([2])
    expect(cellIndices(rows[1][0])).toEqual([3, 4])
    // the continuation cell is numbered in the model even though the editor never renders it
    expect(cellIndices(rows[1][1])).toEqual([5])
    expect(blocks[2].sourceParaIndex).toBe(6)
    expect(blocks[3].sourceParaIndex).toBe(7)
  })

  it('numbers paragraphs of a nested table and of a cell with an anchored text box', async () => {
    const nested = TABLE([TC(P('innerA')) + TC(P('innerB'))])
    const body =
      TABLE(
        [
          TC(P('outer') + nested + '<w:p/>') +
            TC(`<w:p><w:r><w:t>host</w:t></w:r><w:r>${TEXTBOX_ANCHOR}</w:r></w:p>` + P('after')),
        ],
        2,
      ) + P('tail')
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    const blocks = visible(doc.blocks)
    const [outerCell, boxCell] = blocks[0].table!.rows[0]
    // outer 0, innerA 1, innerB 2, the empty paragraph after the nested table 3
    expect(cellIndices(outerCell)).toEqual([0, 3])
    const inner = outerCell.nestedTables![0].rows[0]
    expect(cellIndices(inner[0])).toEqual([1])
    expect(cellIndices(inner[1])).toEqual([2])
    // host 4; the box paragraph is 5 but not modelled; "after" skips to 6
    expect(boxCell.anchoredBoxes).toHaveLength(1)
    expect(cellIndices(boxCell)).toEqual([4, 6])
    expect(blocks[1].sourceParaIndex).toBe(7)
  })

  it('gives a content-control paragraph its own index and numbers an ins-wrapped paragraph', async () => {
    const sdt = (inner: string) =>
      `<w:sdt><w:sdtPr><w:alias w:val="Field"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`
    const ins = (inner: string) =>
      `<w:ins w:id="1" w:author="Author" w:date="2026-01-01T00:00:00Z">${inner}</w:ins>`
    const body =
      P('first') +
      sdt(P('single control')) +
      sdt(P('part one') + P('part two')) +
      ins(P('inserted one') + P('inserted two')) +
      ins(P('inserted alone')) +
      P('last')
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    const at = (text: string) => blockWithText(doc.blocks, text)!.sourceParaIndex
    expect(at('first')).toBe(0)
    expect(at('single control')).toBe(1)
    expect(at('part one')).toBe(2)
    expect(at('part two')).toBe(3)
    // a w:ins wrapping two paragraphs (indices 4 and 5) stamps neither
    expect(at('inserted one')).toBeUndefined()
    expect(at('inserted alone')).toBe(6)
    expect(at('last')).toBe(7)
  })

  it('moves the index of a folded multi-paragraph field onto its displayed tail block', async () => {
    const head =
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r></w:p>'
    const tail =
      '<w:p><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>value</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    const doc = await parseDocx(await buildDocx({ bodyXml: P('before') + head + tail }))
    const blocks = visible(doc.blocks)
    expect(blocks).toHaveLength(3)
    expect(blocks[1].invisibleMarker).toBe(true)
    expect(blocks[1].sourceParaIndex).toBeUndefined()
    expect(blocks[2].fieldDisplay).toBeDefined()
    expect(blocks[2].sourceParaIndex).toBe(1)
  })

  it('does not count a paragraph inside mc:Fallback, a comment or CDATA', async () => {
    const alternate =
      `<mc:AlternateContent ${MC_NS}><mc:Choice Requires="wps">${TRIANGLE_ANCHOR}</mc:Choice>` +
      '<mc:Fallback><w:pict><v:rect xmlns:v="urn:schemas-microsoft-com:vml" style="position:absolute;width:46pt;height:42pt">' +
      '<v:textbox><w:txbxContent><w:p><w:r><w:t>fallback</w:t></w:r></w:p></w:txbxContent></v:textbox>' +
      '</v:rect></w:pict></mc:Fallback></mc:AlternateContent>'
    const body =
      P('opening') +
      '<!-- <w:p><w:r><w:t>commented</w:t></w:r></w:p> -->' +
      `<w:p><w:r><w:t>shape host</w:t></w:r><w:r>${alternate}</w:r></w:p>` +
      '<w:p><w:r><w:t><![CDATA[<w:p/>cdata]]></w:t></w:r></w:p>' +
      P('closing')
    const doc = await parseDocx(await buildDocx({ bodyXml: body }))
    const blocks = visible(doc.blocks)
    expect(blocks).toHaveLength(4)
    expect(blocks.map((b) => b.runs?.map((r) => r.text).join(''))).toEqual([
      'opening',
      'shape host',
      '<w:p/>cdata',
      'closing',
    ])
    // the commented paragraph and the fallback box paragraph are not counted
    expect(blocks.map((b) => b.sourceParaIndex)).toEqual([0, 1, 2, 3])
  })

  it('stamps nothing for an altChunk expansion', async () => {
    const chunk = await buildDocx({ bodyXml: P('chunk one') + P('chunk two') })
    const zip = await JSZip.loadAsync(
      await buildDocx({
        bodyXml: P('before') + '<w:altChunk r:id="rIdChunk"/>' + P('after'),
        extraRels:
          '<Relationship Id="rIdChunk" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="/word/afchunk.dat"/>',
      }),
    )
    zip.file('word/afchunk.dat', chunk)
    const doc = await parseDocx(await zip.generateAsync({ type: 'uint8array' }))
    const blocks = visible(doc.blocks)
    expect(blocks.map((b) => b.runs?.map((r) => r.text).join(''))).toEqual([
      'before',
      'chunk one',
      'chunk two',
      'after',
    ])
    expect(blocks.map((b) => b.sourceParaIndex)).toEqual([0, undefined, undefined, 1])
  })

  it('paragraphStartOffsets: local name p, any prefix, name must end', () => {
    const xml =
      '<w:body><w:p><w:pPr/></w:p><a:p/><w:pict/><!-- <w:p/> --><![CDATA[<w:p/>]]><mc:AlternateContent><mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent><w:p/></w:body>'
    expect(paragraphStartOffsets(xml)).toEqual([
      xml.indexOf('<w:p>'),
      xml.indexOf('<a:p/>'),
      xml.lastIndexOf('<w:p/>'),
    ])
  })

  it('an untouched save is byte-identical; an edited save carries no marker names and no new highlight, shading, bookmark or comment', async () => {
    const bytes = await buildDocx({ bodyXml: TABLE_DOC_BODY + P('edit me') })
    const doc = await parseDocx(bytes)
    const blocks = visible(doc.blocks)
    const untouched = await saveDocx(
      doc,
      blocks.map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
    )
    expect(untouched).toEqual(bytes)

    const target = blockWithText(blocks, 'edit me')!
    expect(target.sourceParaIndex).toBe(8)
    const edited = await saveDocx(
      doc,
      blocks.map((b) =>
        b === target
          ? {
              kind: 'generated' as const,
              block: { type: 'paragraph' as const, runs: [{ text: 'edited text' }] },
            }
          : { kind: 'original' as const, docxIndex: b.docxIndex! },
      ),
    )
    const sourceZip = await JSZip.loadAsync(bytes)
    const savedZip = await JSZip.loadAsync(edited)
    const sourceXml = await sourceZip.file('word/document.xml')!.async('string')
    const savedXml = await savedZip.file('word/document.xml')!.async('string')
    expect(savedXml).toContain('edited text')
    expect(savedXml).not.toContain('sourceParaIndex')
    expect(savedXml).not.toContain('docxPara')
    for (const pattern of [/<w:highlight[\s>/]/g, /<w:shd[\s>/]/g, /<w:bookmarkStart[\s>/]/g]) {
      expect(savedXml.match(pattern)?.length ?? 0).toBe(sourceXml.match(pattern)?.length ?? 0)
    }
    expect(sourceZip.file('word/comments.xml')).toBeNull()
    expect(savedZip.file('word/comments.xml')).toBeNull()
  })
})
