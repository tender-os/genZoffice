import JSZip from 'jszip'
import { describe, expect, test } from 'vitest'
import { parseDocx, patchTableCellTexts, saveDocx } from '../src/index'
import { patchFieldParagraphText } from '../src/text-patch'
import { buildDocx } from './helpers/build-docx'

const formtext =
  '<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Text1"/><w:enabled/><w:textInput/></w:ffData></w:fldChar></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:rPr><w:b/></w:rPr><w:t>(Firma)</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'

const placeholder =
  '<w:p><w:r><w:t>Vergabe-Nr:</w:t></w:r><w:r><w:tab/></w:r>' +
  '<w:sdt><w:sdtPr><w:id w:val="7"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent>' +
  '<w:r><w:rPr><w:rStyle w:val="Platzhaltertext"/></w:rPr><w:t>Klicken Sie hier</w:t></w:r></w:sdtContent></w:sdt></w:p>'

describe('patchFieldParagraphText (Vound fork)', () => {
  test('a text edit inside a FORMTEXT result keeps the field machinery', () => {
    const out = patchFieldParagraphText(formtext, [{ text: 'Vound GmbH' }])
    expect(out).not.toBeNull()
    expect(out).toContain('<w:ffData>')
    expect(out).toContain('FORMTEXT')
    expect(out).toContain('fldCharType="end"')
    expect(out).toContain('<w:rPr><w:b/></w:rPr>')
    expect(out).toContain('>Vound GmbH</w:t>')
    expect(out).not.toContain('(Firma)')
  })
  test('typing into a placeholder control keeps the control and drops the placeholder flag and style', () => {
    const out = patchFieldParagraphText(placeholder, [{ text: 'Vergabe-Nr:\t2026-1' }])
    expect(out).not.toBeNull()
    expect(out).toContain('<w:sdt>')
    expect(out).toContain('<w:tab/>')
    expect(out).toContain('>2026-1</w:t>')
    expect(out).not.toContain('showingPlcHdr')
    expect(out).not.toContain('Platzhaltertext')
  })
  test('two controls of one paragraph edited before one save each get their own text', () => {
    const control = (id: string, text: string) =>
      `<w:sdt><w:sdtPr><w:id w:val="${id}"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>${text}</w:t></w:r></w:sdtContent></w:sdt>`
    const two =
      `<w:p><w:r><w:t xml:space="preserve">Ort: </w:t></w:r>${control('1', 'Ort eintragen')}` +
      `<w:r><w:t xml:space="preserve">, Datum: </w:t></w:r>${control('2', 'Datum eingeben')}</w:p>`
    const out = patchFieldParagraphText(two, [{ text: 'Ort: Berlin, Datum: 09.11.2026' }])!
    expect(out).toContain('<w:r><w:t xml:space="preserve">Ort: </w:t></w:r>')
    expect(out).toContain('<w:sdtContent><w:r><w:t xml:space="preserve">Berlin</w:t></w:r></w:sdtContent>')
    expect(out).toContain('<w:r><w:t xml:space="preserve">, Datum: </w:t></w:r>')
    expect(out).toContain('<w:sdtContent><w:r><w:t xml:space="preserve">09.11.2026</w:t></w:r></w:sdtContent>')
    expect(out).not.toContain('showingPlcHdr')
  })
  test('a line break typed into a control is written as a w:br inside it', () => {
    const control =
      '<w:p><w:sdt><w:sdtPr><w:id w:val="8"/></w:sdtPr><w:sdtContent><w:r><w:t>Erste</w:t></w:r></w:sdtContent></w:sdt></w:p>'
    const once = patchFieldParagraphText(control, [{ text: 'Erste\nZweite' }])!
    expect(once).toContain(
      '<w:sdtContent><w:r><w:t xml:space="preserve">Erste</w:t><w:br/><w:t xml:space="preserve">Zweite</w:t></w:r></w:sdtContent>',
    )
    const twice = patchFieldParagraphText(once, [{ text: 'Erste\nZweite\nDritte' }])!
    expect(twice).toContain(
      '<w:t xml:space="preserve">Erste</w:t><w:br/><w:t xml:space="preserve">Zweite</w:t><w:br/><w:t xml:space="preserve">Dritte</w:t>',
    )
    expect(patchFieldParagraphText(twice, [{ text: 'Erste\nZweiteDritte' }])).not.toContain(
      'Zweite</w:t><w:br/>',
    )
  })
  test('a control whose text did not change keeps its placeholder flag', () => {
    // the edit sits inside the first run; an insertion exactly at the control boundary is
    // ambiguous and may be attributed to the control (then the flag goes, as if typed into it)
    const out = patchFieldParagraphText(placeholder, [{ text: 'Vergabe NEU-Nr:\tKlicken Sie hier' }])
    expect(out).not.toBeNull()
    expect(out).toContain('showingPlcHdr')
  })
  test('falls back when a tab was added around a plain field, or when there is no field', () => {
    const pageField =
      '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    expect(patchFieldParagraphText(pageField, [{ text: '\t1' }])).toBeNull()
    expect(patchFieldParagraphText('<w:p><w:r><w:t>plain</w:t></w:r></w:p>', [{ text: 'x' }])).toBeNull()
  })
  test('a form field keeps its paragraph when a line break or tab is added: each becomes a space', () => {
    const out = patchFieldParagraphText(formtext, [{ text: 'Zeile 1\nZeile\t2' }])
    expect(out).toContain('<w:ffData>')
    expect(out).toContain('<w:t xml:space="preserve">Zeile 1 Zeile 2</w:t>')
  })
  test('a control keeps its paragraph when its tab is removed: the tab stays', () => {
    const out = patchFieldParagraphText(placeholder, [{ text: 'Vergabe-Nr:2026' }])
    expect(out).toContain('<w:sdt>')
    expect(out).toContain('<w:tab/>')
  })
  test('falls back for a checkbox glyph that has no w:t to map onto', () => {
    const checkbox =
      '<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:checkBox><w:default w:val="0"/></w:checkBox></w:ffData></w:fldChar></w:r>' +
      '<w:r><w:instrText> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    expect(patchFieldParagraphText(checkbox, [{ text: '☒' }])).toBeNull()
  })
})

const NOTE_REF = '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r>'
const CONTROL = '<w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent><w:r><w:t>Box</w:t></w:r></w:sdtContent></w:sdt>'
const LABEL = '<w:r><w:t>Option A</w:t></w:r>'
const legacyBox = (state: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Box"/><w:enabled/><w:checkBox><w:sizeAuto/>${state}</w:checkBox></w:ffData></w:fldChar></w:r>` +
  '<w:r><w:instrText> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
const CHECKED_BOX = legacyBox('<w:default w:val="1"/>')
const CHECKBOX_CONTROL =
  '<w:sdt><w:sdtPr><w:id w:val="5"/><w14:checkbox xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w14:checked w14:val="1"/></w14:checkbox></w:sdtPr>' +
  '<w:sdtContent><w:r><w:t>☒</w:t></w:r></w:sdtContent></w:sdt>'

const FOOTNOTES = {
  path: 'word/footnotes.xml',
  contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  xml:
    '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> note</w:t></w:r></w:p></w:footnote></w:footnotes>',
}
const FOOTNOTES_REL =
  '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'

/** types one character at the end of the paragraph and saves, as an auto-save does */
async function typeAndSave(bytes: Uint8Array, typed: string): Promise<Uint8Array> {
  const doc = await parseDocx(bytes)
  const block = doc.blocks[0]
  const runs = block.runs ?? []
  const edited = runs.map((run, i) => (i === runs.length - 1 ? { ...run, text: run.text + typed } : run))
  return saveDocx(doc, [{ kind: 'generated', block: { type: 'paragraph', runs: edited }, docxIndex: 0 }])
}

const documentXml = async (bytes: Uint8Array) =>
  (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')

const savedTwice = async (paragraph: string) => {
  const built = await buildDocx({ bodyXml: paragraph, extraParts: [FOOTNOTES], extraRels: FOOTNOTES_REL })
  return documentXml(await typeAndSave(await typeAndSave(built, 'x'), 'y'))
}

describe('text the run model holds but the paragraph does not keep in w:t (Vound fork)', () => {
  test('two saves of an edited paragraph with a footnote reference add nothing but the typed text', async () => {
    const xml = await savedTwice(`<w:p>${CHECKBOX_CONTROL}${NOTE_REF}<w:r><w:tab/></w:r>${LABEL}</w:p>`)
    expect(xml).toContain('>Option Axy</w:t>')
    expect(xml.match(/<w:footnoteReference /g)).toHaveLength(1)
    expect(xml.match(/<w:tab\/>/g)).toHaveLength(1)
    expect(xml).toContain('<w:sdt>')
    expect(xml).toContain('>☒</w:t>')
  })

  test.each([
    ['a symbol', '<w:r><w:sym w:font="Wingdings" w:char="F078"/></w:r>'],
    ['a formula', '<m:oMath><m:r><m:t>ab</m:t></m:r></m:oMath>'],
    ['a tracked deletion', '<w:del w:id="9" w:author="a"><w:r><w:delText>gone</w:delText></w:r></w:del>'],
  ])('two saves of a paragraph with a control and %s add nothing but the typed text', async (_, atom) => {
    const xml = await savedTwice(`<w:p>${CONTROL}${atom}${LABEL}</w:p>`)
    expect(xml).toContain('>Option Axy</w:t>')
    expect(xml).toContain(atom)
  })

  test('two saves of a labelled legacy checkbox keep its glyph out of the label', async () => {
    const xml = await savedTwice(`<w:p>${CHECKED_BOX}${NOTE_REF}<w:r><w:tab/></w:r>${LABEL}</w:p>`)
    expect(xml).toContain('>Option Axy</w:t>')
    expect(xml).toContain('<w:ffData>')
    expect(xml).not.toContain('☒')
  })

  test('a footnote reference in a table cell paragraph adds nothing', () => {
    const cell = (paragraph: string) =>
      `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${paragraph}</w:tc></w:tr></w:tbl>`
    const out = patchTableCellTexts(cell(`<w:p>${CONTROL}${NOTE_REF}${LABEL}</w:p>`), [
      [[{ runs: [{ text: 'Box' }, { text: '1', noteRef: { kind: 'footnote', id: '2' } }, { text: 'Option Ax' }] }]],
    ])
    expect(out).toContain('>Option Ax</w:t>')
    expect(out).toContain('<w:sdt>')
  })

  test('a ticked or deleted legacy checkbox falls back to the rebuild, which saves its state', () => {
    const paragraph = `<w:p>${legacyBox('<w:default w:val="0"/>')}${LABEL}</w:p>`
    const box = { fldBeginXml: '<w:r/>' }
    const unchanged = [{ text: '☐', ...box }, { text: 'Option Ax' }]
    expect(patchFieldParagraphText(paragraph, unchanged)).toContain('>Option Ax</w:t>')
    expect(patchFieldParagraphText(paragraph, [{ text: '☒', ...box }, { text: 'Option A' }])).toBeNull()
    expect(patchFieldParagraphText(paragraph, [{ text: 'Option A' }])).toBeNull()
  })

  test('a saved tick of a labelled legacy checkbox lands in w:checked, not in the label', async () => {
    const built = await buildDocx({ bodyXml: `<w:p>${legacyBox('<w:default w:val="0"/>')}${LABEL}</w:p>` })
    const doc = await parseDocx(built)
    const runs = doc.blocks[0].runs!.map((run) => (run.fldBeginXml ? { ...run, text: '☒' } : run))
    const xml = await documentXml(
      await saveDocx(doc, [{ kind: 'generated', block: { type: 'paragraph', runs }, docxIndex: 0 }]),
    )
    expect(xml).toContain('<w:checked w:val="1"/>')
    expect(xml).toContain('>Option A</w:t>')
    expect(xml).not.toContain('☒')
  })

  test.each([
    ['a removed footnote reference', `<w:p>${CONTROL}${NOTE_REF}${LABEL}</w:p>`, [{ text: 'Box' }, { text: 'Option A' }]],
    ['an added footnote reference', `<w:p>${CONTROL}${LABEL}</w:p>`, [{ text: 'Box' }, { text: '1', noteRef: { kind: 'footnote' as const, id: '2' } }, { text: 'Option A' }]],
    ['a removed tracked deletion', `<w:p>${CONTROL}<w:del w:id="9" w:author="a"><w:r><w:delText>gone</w:delText></w:r></w:del>${LABEL}</w:p>`, [{ text: 'Box' }, { text: 'Option A' }]],
  ])('falls back for %s, which a w:t patch cannot save', (_, paragraph, runs) => {
    expect(patchFieldParagraphText(paragraph, runs)).toBeNull()
  })

  test.each([
    ['a no-break hyphen', '<w:r><w:t>Pre</w:t><w:noBreakHyphen/><w:t>fix</w:t></w:r>'],
    ['a soft hyphen', '<w:r><w:t>Pre</w:t><w:softHyphen/><w:t>fix</w:t></w:r>'],
    ['a phonetic guide', '<w:r><w:ruby><w:rubyPr/><w:rt><w:r><w:t>rt</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>base</w:t></w:r></w:rubyBase></w:ruby></w:r>'],
  ])('falls back for %s, whose text does not map onto w:t', (_, inline) => {
    expect(patchFieldParagraphText(`<w:p>${CONTROL}${inline}${LABEL}</w:p>`, [{ text: 'x' }])).toBeNull()
  })
})
