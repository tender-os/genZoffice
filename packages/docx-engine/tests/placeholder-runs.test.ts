import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

// Placeholder text (Vound fork): what a content control shows while it holds no content (w:showingPlcHdr), and
// a legacy text form field's default text, are marked so the review editor can select them whole. A run that
// only looks like a hint (a grey character style, no control around it) is the document's own text.
const PLACEHOLDER_STYLE =
  '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:rPr><w:color w:val="808080"/></w:rPr></w:style>'
const run = (text: string, rPr = '') => `<w:r>${rPr && `<w:rPr>${rPr}</w:rPr>`}<w:t xml:space="preserve">${text}</w:t></w:r>`
const hint = (text: string) => run(text, '<w:rStyle w:val="PlaceholderText"/>')
const sdtPr = (flag: string) => `<w:sdtPr><w:id w:val="7"/>${flag}<w:text/></w:sdtPr>`
const inlineControl = (flag: string, content: string) => `<w:sdt>${sdtPr(flag)}<w:sdtContent>${content}</w:sdtContent></w:sdt>`
const p = (inner: string) => `<w:p>${inner}</w:p>`
const formText = (defaultText: string, result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="F"/><w:enabled/><w:textInput><w:default w:val="${defaultText}"/></w:textInput></w:ffData></w:fldChar></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${run(result)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`
const table = (...cells: string[]) =>
  `<w:tbl><w:tblGrid>${cells.map(() => '<w:gridCol w:w="3000"/>').join('')}</w:tblGrid><w:tr>${cells.map((c) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${c}</w:tc>`).join('')}</w:tr></w:tbl>`

type ParsedRun = { text: string; placeholder?: true }
/** text and placeholder flag of each run */
const flags = (runs: ParsedRun[] | undefined) => (runs ?? []).map((r) => [r.text, r.placeholder === true])
const parse = async (bodyXml: string) => parseDocx(await buildDocx({ bodyXml, extraStylesXml: PLACEHOLDER_STYLE }))

describe('placeholder runs', () => {
  it('marks the text of a control showing its placeholder, not the text around it or of a filled control', async () => {
    const doc = await parse(
      p(run('Firma: ') + inlineControl('<w:showingPlcHdr/>', hint('Name eingeben'))) +
        p(inlineControl('', run('Muster GmbH'))) +
        p(inlineControl('<w:showingPlcHdr w:val="0"/>', run('Beispiel AG'))) +
        p(hint('Sitz des Unternehmens')) +
        p(run('Ort: ') + inlineControl('<w:showingPlcHdr/>', run('Ort eingeben'))),
    )
    expect(doc.blocks.slice(0, 5).map((b) => flags(b.runs))).toEqual([
      [['Firma: ', false], ['Name eingeben', true]],
      [['Muster GmbH', false]],
      [['Beispiel AG', false]],
      [['Sitz des Unternehmens', false]],
      [['Ort: ', false], ['Ort eingeben', true]],
    ])
  })

  it('marks the paragraph of a block-level control showing its placeholder', async () => {
    const doc = await parse(`<w:sdt>${sdtPr('<w:showingPlcHdr/>')}<w:sdtContent>${p(hint('Rechtsform'))}</w:sdtContent></w:sdt>`)
    expect(flags(doc.blocks[0].runs)).toEqual([['Rechtsform', true]])
  })

  it('marks placeholders in table cells: inline and block-level controls, and a form field showing its default text', async () => {
    const doc = await parse(
      table(
        p(inlineControl('<w:showingPlcHdr/>', hint('USt-IdNr.'))),
        `<w:sdt>${sdtPr('<w:showingPlcHdr/>')}<w:sdtContent>${p(hint('Rechtsform'))}</w:sdtContent></w:sdt>`,
        p(formText('Ort', 'Ort')),
        p(formText('Ort', 'Berlin')),
        p(hint('Sitz des Unternehmens')),
      ),
    )
    const cells = doc.blocks[0].table!.rows[0]
    expect(cells.map((cell) => flags(cell.richParas?.[0]?.runs))).toEqual([
      [['USt-IdNr.', true]],
      [['Rechtsform', true]],
      [['Ort', true]],
      [['Berlin', false]],
      [['Sitz des Unternehmens', false]],
    ])
  })
})
