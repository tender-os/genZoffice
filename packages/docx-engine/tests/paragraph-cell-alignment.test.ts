import { describe, expect, test } from 'vitest'
import { withCellVerticalAlignment, withParagraphAlignment } from '../src/paragraph-cell-alignment'
import { withRestyledRuns } from '../src/restyled-runs'

describe('withParagraphAlignment (Vound fork)', () => {
  test('sets w:jc in CT_PPr order, replacing one there was', () => {
    expect(withParagraphAlignment('<w:p><w:r><w:t>a</w:t></w:r></w:p>', 'center')).toBe(
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>',
    )
    expect(
      withParagraphAlignment(
        '<w:p><w:pPr><w:jc w:val="left"/><w:rPr><w:b/></w:rPr></w:pPr></w:p>',
        'right',
      ),
    ).toBe('<w:p><w:pPr><w:jc w:val="right"/><w:rPr><w:b/></w:rPr></w:pPr></w:p>')
    expect(
      withParagraphAlignment(
        '<w:p w14:paraId="1"><w:pPr><w:ind w:left="0"/><w:rPr/></w:pPr></w:p>',
        'justify',
      ),
    ).toBe(
      '<w:p w14:paraId="1"><w:pPr><w:ind w:left="0"/><w:jc w:val="both"/><w:rPr/></w:pPr></w:p>',
    )
  })
})

describe('withCellVerticalAlignment (Vound fork)', () => {
  test('sets w:vAlign in the cell properties, after the width and before hideMark', () => {
    expect(
      withCellVerticalAlignment(
        '<w:tc><w:tcPr><w:tcW w:w="1"/><w:hideMark/></w:tcPr><w:p/></w:tc>',
        'bottom',
      ),
    ).toBe(
      '<w:tc><w:tcPr><w:tcW w:w="1"/><w:vAlign w:val="bottom"/><w:hideMark/></w:tcPr><w:p/></w:tc>',
    )
    expect(withCellVerticalAlignment('<w:tc><w:p/></w:tc>', 'center')).toBe(
      '<w:tc><w:tcPr><w:vAlign w:val="center"/></w:tcPr><w:p/></w:tc>',
    )
  })
})

describe('withRestyledRuns size override (Vound fork)', () => {
  test('a run the editor took the size away from loses w:sz and w:szCs', () => {
    const RPR = '<w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
    const para = `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r>${RPR}<w:t>Berlin</w:t></w:r></w:p>`
    expect(withRestyledRuns(para, [{ text: 'Berlin', rawRPr: RPR }])).toContain(
      '<w:r><w:rPr><w:b/></w:rPr><w:t>Berlin</w:t></w:r>',
    )
    expect(withRestyledRuns(para, [{ text: 'Berlin', sizeHalfPoints: null }])).toContain(
      '<w:r><w:rPr><w:b/></w:rPr><w:t>Berlin</w:t></w:r>',
    )
    expect(withRestyledRuns(para, [{ text: 'Berlin' }])).toBe(para)
  })
})
