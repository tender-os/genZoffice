import { describe, expect, it } from 'vitest'
import { patchFieldParagraphXml } from '../src/index'
import { withoutFilledFieldHighlight } from '../src/form-text-highlight'
import { patchFieldParagraphText } from '../src/text-patch'

const MARKED = '<w:rPr><w:highlight w:val="yellow"/></w:rPr>'
const marked = (inner: string) => `<w:r>${MARKED}${inner}</w:r>`
const DEFAULT_TEXT = '[Bitte Name angeben]'
const EN_SPACES = ' '.repeat(5)

const field = (result: string | null, defaultText = DEFAULT_TEXT) =>
  marked(
    `<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val=""/><w:enabled/><w:textInput><w:default w:val="${defaultText}"/></w:textInput></w:ffData></w:fldChar>`,
  ) +
  marked('<w:instrText xml:space="preserve"> FORMTEXT </w:instrText>') +
  marked('') +
  marked('<w:fldChar w:fldCharType="separate"/>') +
  (result === null ? '' : marked(`<w:t xml:space="preserve">${result}</w:t>`)) +
  marked('<w:fldChar w:fldCharType="end"/>')

const label = (text: string) => marked(`<w:t xml:space="preserve">${text}</w:t>`)
const para = (inner: string) => `<w:p>${inner}</w:p>`
const highlights = (xml: string) => xml.split('<w:highlight').length - 1

const RESULTS = ['Erika Muster', EN_SPACES, DEFAULT_TEXT, '', null] as const
const OPEN_RESULTS = [EN_SPACES, DEFAULT_TEXT, '', null]

describe('a legacy text field the author marked with highlight (Vound fork)', () => {
  it('loses the highlight on every run, begin to end, once it holds a value', () => {
    const out = withoutFilledFieldHighlight(para(`${label('Mitglied ')}${field('Erika Muster')}`))
    expect(out).toBe(
      para(`${label('Mitglied ')}${field('Erika Muster').replaceAll(MARKED, '<w:rPr></w:rPr>')}`),
    )
  })

  it.each(OPEN_RESULTS)('keeps it while open (result %j)', (result) => {
    const xml = para(field(result))
    expect(withoutFilledFieldHighlight(xml)).toBe(xml)
  })

  it('strips exactly the fields holding a value and leaves every other byte as it was, for each mix', () => {
    for (const first of RESULTS) {
      for (const second of RESULTS) {
        const xml = para(`${label('A ')}${field(first)}${label(' / ')}${field(second)}`)
        const expected = para(
          `${label('A ')}${first === RESULTS[0] ? field(first).replaceAll(MARKED, '<w:rPr></w:rPr>') : field(first)}` +
            `${label(' / ')}${second === RESULTS[0] ? field(second).replaceAll(MARKED, '<w:rPr></w:rPr>') : field(second)}`,
        )
        expect(withoutFilledFieldHighlight(xml)).toBe(expected)
      }
    }
  })

  it('saves a typed value without it through the paragraph patch, the label keeping its own', () => {
    const xml = para(`${label('Mitglied ')}${field(DEFAULT_TEXT)}`)
    const out = patchFieldParagraphText(xml, [
      { text: 'Mitglied ' },
      { text: 'Erika Muster', formText: 0 },
    ])
    expect(out).toContain('>Erika Muster<')
    expect(highlights(out ?? '')).toBe(1)
    expect(out).toContain(label('Mitglied '))
  })

  it('keeps it when an undo took the value back to the default text', () => {
    const xml = para(`${label('Mitglied ')}${field(DEFAULT_TEXT)}`)
    const out = patchFieldParagraphText(xml, [
      { text: 'Mitglied ' },
      { text: DEFAULT_TEXT, formText: 0, placeholder: true },
    ])
    expect(highlights(out ?? '')).toBe(highlights(xml))
  })

  it('saves a field passthrough value without it', () => {
    const out = patchFieldParagraphXml(para(field(EN_SPACES)), { left: 'Erika Muster' })
    expect(out).toContain('>Erika Muster<')
    expect(highlights(out)).toBe(0)
  })
})

describe('a settled field passthrough (Vound fork)', () => {
  it('empties each result still showing its default text and keeps values and blanks', () => {
    const xml = para(
      `${field(DEFAULT_TEXT)}${label(' / ')}${field('Erika Muster')}${field(EN_SPACES)}`,
    )
    const out = patchFieldParagraphXml(xml, { settled: true })
    const results = out
      .split('w:fldCharType="separate"/>')
      .slice(1)
      .map((part) => part.split('<w:fldChar w:fldCharType="end"')[0].replace(/<[^>]+>/g, ''))
    expect(results).toEqual(['', 'Erika Muster', EN_SPACES])
    expect(highlights(out)).toBe(1)
  })

  it('keeps the default text of a passthrough the host did not settle', () => {
    expect(patchFieldParagraphXml(para(field(DEFAULT_TEXT)), {})).toContain(`>${DEFAULT_TEXT}<`)
  })
})
