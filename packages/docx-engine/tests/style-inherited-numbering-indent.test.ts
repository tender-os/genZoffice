import { describe, expect, it } from 'vitest'
import { buildAnchoredTextboxParagraphXml } from '../src/generate'
import { parseDocx } from '../src/index'
import type { EffectivePPr } from '../src/resolve-props'
import { buildStyleNumberingDocx, DIRECT_NUMPR } from './helpers/style-numbering-docx'

const para = (pPr: string, text: string) =>
  `<w:p><w:pPr>${pPr}<w:jc w:val="left"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

async function effectiveIndentOf(paragraph: string) {
  const doc = await parseDocx(await buildStyleNumberingDocx(paragraph))
  const block = doc.blocks[0]
  expect(block.type).toBe('listItem')
  return (block.format as { effective: EffectivePPr }).effective.indent
}

describe('numbering inherited from the paragraph style', () => {
  it('lets the style chain indent win over the level indent', async () => {
    expect(await effectiveIndentOf(para('<w:pStyle w:val="Bar"/>', 'Heading'))).toEqual({
      left: 499,
      right: -144,
      firstLine: -357,
    })
  })

  it('keeps the level indent when the paragraph has its own w:numPr', async () => {
    expect(
      await effectiveIndentOf(para(`<w:pStyle w:val="Bar"/>${DIRECT_NUMPR}`, 'Heading')),
    ).toEqual({ left: 9999, right: -144, firstLine: -360 })
  })

  it('keeps the level indent over an indent from above the numbering style', async () => {
    expect(await effectiveIndentOf(para('<w:pStyle w:val="Plain"/>', 'Item'))).toEqual({
      left: 9999,
      right: 0,
      firstLine: -360,
    })
  })

  it('marks the list block as numbered through its style', async () => {
    const doc = await parseDocx(
      await buildStyleNumberingDocx(
        para('<w:pStyle w:val="Bar"/>', 'Heading') +
          para(`<w:pStyle w:val="Bar"/>${DIRECT_NUMPR}`, 'Heading'),
      ),
    )
    expect(doc.blocks.slice(0, 2).map((b) => b.listFromStyle)).toEqual([true, undefined])
    expect(doc.styles.get('Bar')?.listIndent).toEqual({
      indentLeftTwips: 499,
      indentRightTwips: -144,
      indentFirstLineTwips: -357,
    })
    expect(doc.styles.get('Plain')?.listIndent).toEqual({})
    expect(doc.styles.get('Normal')?.listIndent).toBeUndefined()
  })
})

function textboxWith(pPrs: string[]): string {
  const xml = buildAnchoredTextboxParagraphXml({
    anchor: 'paragraph',
    xEmu: 0,
    yEmu: 0,
    widthEmu: 3000000,
    heightEmu: 1000000,
    insetsEmu: { l: 0, t: 0, r: 0, b: 0 },
    zOrder: 1,
    id: 7,
    paragraphs: pPrs.map((_, i) => ({ runs: [{ text: `item ${i + 1}` }] })),
    holderLineTwips: 20,
  })
  let i = 0
  return xml.replace(/<w:p><w:r>/g, () => `<w:p><w:pPr>${pPrs[i++]}</w:pPr><w:r>`)
}

describe('textbox list markers numbered through the paragraph style', () => {
  it('take the style indent, and the level indent for a direct w:numPr', async () => {
    const doc = await parseDocx(
      await buildStyleNumberingDocx(
        textboxWith(['<w:pStyle w:val="Bar"/>', `<w:pStyle w:val="Bar"/>${DIRECT_NUMPR}`]),
      ),
    )
    const [viaStyle, direct] = doc.blocks.flatMap((b) => b.textboxes ?? [])[0]!.paras
    expect(viaStyle.listMarker).toMatchObject({ indentLeft: 499, indentRight: -144, hanging: 357 })
    expect(direct.listMarker).toMatchObject({ indentLeft: 9999, hanging: 360 })
    expect(direct.listMarker?.indentRight).toBeUndefined()
  })
})
