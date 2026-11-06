import { describe, expect, it } from 'vitest'
import { parseDocx, patchFieldParagraphXml } from '../src/index'
import { buildDocx } from './helpers/build-docx'

// An empty legacy text form field as Word writes it: its result is five en spaces, one per run.
const BLANK_FORMTEXT =
  '<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Text2"/><w:enabled/><w:textInput/></w:ffData></w:fldChar></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>\u2002</w:t></w:r><w:r><w:t>\u2002</w:t></w:r><w:r><w:t>\u2002</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
// Forms draw a long name line as several such fields side by side in one paragraph.
const BLANK_FIELD_RUNS = BLANK_FORMTEXT.slice('<w:p>'.length, -'</w:p>'.length)
const BLANK_NAME_LINE = `<w:p>${BLANK_FIELD_RUNS.repeat(3)}</w:p>`

/** the text of each field result in the paragraph, in order */
const fieldResults = (xml: string) =>
  xml
    .split('fldCharType="separate"/>')
    .slice(1)
    .map((part) =>
      part.slice(0, part.indexOf('<w:fldChar w:fldCharType="end"')).replace(/<[^>]+>/g, ''),
    )

describe('a blank legacy text form field (Vound fork)', () => {
  it('parses as a text field with an empty result on its empty line, so it can be typed into', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BLANK_FORMTEXT }))
    expect(doc.blocks[0].fieldDisplay).toMatchObject({ kind: 'text', left: '', emptyLine: true })
    // unlike a lone field end, it has result text to take a value
    expect(doc.blocks[0].fieldDisplay?.fieldEndOnly).toBeUndefined()
  })

  it('takes a typed result inside the field result and keeps the field machinery', () => {
    const out = patchFieldParagraphXml(BLANK_FORMTEXT, { left: 'Muster GmbH' })
    const separate = 'fldCharType="separate"/>'
    const result = out.slice(
      out.indexOf(separate) + separate.length,
      out.indexOf('<w:fldChar w:fldCharType="end"'),
    )
    expect(result.replace(/<[^>]+>/g, '')).toBe('Muster GmbH')
    expect(out).toContain('<w:ffData>')
    expect(out).toContain(' FORMTEXT ')
  })

  it('leaves a blank result with a line break, a blank non-form field, and text runs outside the field as they were', async () => {
    const withBreak = BLANK_FORMTEXT.replace(
      '<w:r><w:fldChar w:fldCharType="end"/>',
      '<w:r><w:br/></w:r><w:r><w:fldChar w:fldCharType="end"/>',
    )
    const page = BLANK_FORMTEXT.replace(' FORMTEXT ', ' PAGE ')
    // a save spreads the new result over every w:t of the paragraph (patchFieldParagraphXml): a space run
    // before or after the field would take part of the typed text outside the field result
    const space = '<w:r><w:t xml:space="preserve"> </w:t></w:r>'
    const trailingSpace = BLANK_FORMTEXT.replace('</w:p>', `${space}</w:p>`)
    const leadingSpace = BLANK_FORMTEXT.replace('<w:p>', `<w:p>${space}`)
    for (const bodyXml of [withBreak, page, trailingSpace, leadingSpace]) {
      const doc = await parseDocx(await buildDocx({ bodyXml }))
      expect(doc.blocks[0].fieldDisplay?.emptyLine).toBeUndefined()
    }
  })

  it('parses a line of several blank fields as one empty text field', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BLANK_NAME_LINE }))
    expect(doc.blocks[0].fieldDisplay).toMatchObject({ kind: 'text', left: '', emptyLine: true })
  })

  it("writes a line of several fields into the first field's result and leaves the other fields as they were", () => {
    const enSpaces = '\u2002'.repeat(3)
    const filled = patchFieldParagraphXml(BLANK_NAME_LINE, { left: 'Erika Mustermann' })
    expect(fieldResults(filled)).toEqual(['Erika Mustermann', enSpaces, enSpaces])
    expect(filled.match(/ FORMTEXT /g)).toHaveLength(3)
    // the next edit finds the value where the last one put it
    expect(fieldResults(patchFieldParagraphXml(filled, { left: 'Max Muster' }))).toEqual([
      'Max Muster',
      enSpaces,
      enSpaces,
    ])
  })
})
