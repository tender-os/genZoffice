import { describe, expect, test } from 'vitest'
import { withFormTextStyles } from '../src/form-text-style'
import { mergeRPrModel } from '../src/generate'
import type { Run } from '../src/types'

const RPR = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>'
const field = (result: string) =>
  '<w:p><w:r><w:t xml:space="preserve">Ort: </w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="F"/><w:textInput/></w:ffData></w:fldChar></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `${result}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`
const PARA = field(`<w:r>${RPR}<w:t>Berlin</w:t></w:r>`)
const rPrOf = (run: Run) => mergeRPrModel(run.rawRPr ?? '', run, false)
const label: Run = { text: 'Ort: ' }
const city = (patch: Partial<Run> = {}, text = 'Berlin'): Run => ({
  text,
  formText: 0,
  rawRPr: RPR,
  fontAscii: 'Calibri',
  font: 'Calibri',
  sizeHalfPoints: 22,
  ...patch,
})

describe('withFormTextStyles (Vound fork)', () => {
  test('a result whose runs kept their formatting keeps its bytes', () => {
    expect(withFormTextStyles(PARA, [label, city()], rPrOf)).toBe(PARA)
  })

  test('a new face and size of the field reach its result runs', () => {
    const out = withFormTextStyles(
      PARA,
      [label, city({ fontAscii: 'Arial', sizeHalfPoints: 24 })],
      rPrOf,
    )
    expect(out).toContain('w:ascii="Arial"')
    expect(out).toContain('<w:sz w:val="24"/>')
    expect(out).toContain('>Berlin</w:t>')
    expect(out).toContain('FORMTEXT')
    expect(out.startsWith('<w:p><w:r><w:t xml:space="preserve">Ort: </w:t></w:r>')).toBe(true)
  })

  test('part of the value restyled splits the result into one run per formatting', () => {
    const out = withFormTextStyles(
      PARA,
      [label, city({}, 'Ber'), city({ sizeHalfPoints: 28 }, 'lin')],
      rPrOf,
    )
    expect(out).toContain(`<w:r>${RPR}<w:t xml:space="preserve">Ber</w:t></w:r>`)
    expect(out).toMatch(/<w:sz w:val="28"\/>.*>lin<\/w:t>/)
  })

  test('a paragraph without a legacy text field is left alone', () => {
    const plain = `<w:p><w:r>${RPR}<w:t>Berlin</w:t></w:r></w:p>`
    expect(withFormTextStyles(plain, [city({ fontAscii: 'Arial' })], rPrOf)).toBe(plain)
  })
})
