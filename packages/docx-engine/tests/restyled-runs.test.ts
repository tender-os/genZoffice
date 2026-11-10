import { describe, expect, test } from 'vitest'
import { withRestyledRuns } from '../src/restyled-runs'
import type { Run } from '../src/types'

const RPR = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>'
const formText = (result: string) =>
  '<w:p><w:r><w:t xml:space="preserve">Ort: </w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="F"/><w:textInput/></w:ffData></w:fldChar></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `${result}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`
const PARA = formText(`<w:r>${RPR}<w:t>Berlin</w:t></w:r>`)
const CONTROL =
  '<w:p><w:r><w:t xml:space="preserve">Ort: </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="1"/><w:text/></w:sdtPr>' +
  `<w:sdtContent><w:r>${RPR}<w:t>Berlin</w:t></w:r></w:sdtContent></w:sdt></w:p>`
const label: Run = { text: 'Ort: ' }
const city = (patch: Partial<Run> = {}, text = 'Berlin'): Run => ({
  text,
  formText: 0,
  fontAscii: 'Calibri',
  sizeHalfPoints: 22,
  ...patch,
})

describe('withRestyledRuns (Vound fork)', () => {
  test('runs that kept their face and size keep their bytes', () => {
    expect(withRestyledRuns(PARA, [label, city()])).toBe(PARA)
    expect(withRestyledRuns(CONTROL, [label, city()])).toBe(CONTROL)
  })

  test('a new face and size reach the field result run, and nothing else changes', () => {
    const out = withRestyledRuns(PARA, [label, city({ fontAscii: 'Arial', sizeHalfPoints: 24 })])
    expect(out).toBe(
      PARA.replace(
        RPR,
        '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>',
      ),
    )
  })

  test('a new face reaches the run inside a content control, and the control stays', () => {
    const out = withRestyledRuns(CONTROL, [label, city({ fontAscii: 'Arial' })])
    expect(out).toContain(
      '<w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>Berlin</w:t></w:r></w:sdtContent>',
    )
  })

  test('part of a run restyled splits it into one run per formatting', () => {
    const out = withRestyledRuns(PARA, [
      label,
      city({}, 'Ber'),
      city({ sizeHalfPoints: 28 }, 'lin'),
    ])
    expect(out).toContain(`<w:r>${RPR}<w:t xml:space="preserve">Ber</w:t></w:r>`)
    expect(out).toContain(
      '<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t xml:space="preserve">lin</w:t></w:r>',
    )
  })

  test('a run without properties gets them, after a theme face is replaced', () => {
    const bare = formText('<w:r><w:t>Berlin</w:t></w:r>')
    expect(
      withRestyledRuns(bare, [label, city({ fontAscii: 'Arial', sizeHalfPoints: undefined })]),
    ).toContain(
      '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>Berlin</w:t></w:r>',
    )
    const themed = formText(
      '<w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/></w:rPr><w:t>Berlin</w:t></w:r>',
    )
    expect(
      withRestyledRuns(themed, [label, city({ fontAscii: 'Arial', sizeHalfPoints: undefined })]),
    ).toContain('<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>')
  })

  test('text that does not line up with the runs leaves the paragraph alone', () => {
    expect(withRestyledRuns(PARA, [label, city({ fontAscii: 'Arial' }, 'Köln')])).toBe(PARA)
  })
})
