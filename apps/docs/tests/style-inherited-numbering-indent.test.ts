import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import {
  buildStyleNumberingDocx,
  DIRECT_NUMPR,
} from '../../../packages/docx-engine/tests/helpers/style-numbering-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const VIA_STYLE = '<w:p><w:pPr><w:pStyle w:val="Bar"/></w:pPr><w:r><w:t>Via style</w:t></w:r></w:p>'
const DIRECT = `<w:p><w:pPr><w:pStyle w:val="Bar"/>${DIRECT_NUMPR}</w:pPr><w:r><w:t>Direct</w:t></w:r></w:p>`
const BODY = '<w:p><w:r><w:t>Body</w:t></w:r></w:p>'

async function openEditor() {
  const parsed = await parseDocx(await buildStyleNumberingDocx(VIA_STYLE + DIRECT + BODY))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.storage.listNumbering.styles = parsed.styles
  editor.storage.listNumbering.defs = parsed.numbering
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  return { parsed, editor }
}

const listStyles = (editor: Editor) =>
  Array.from(editor.view.dom.querySelectorAll('.doc-li')).map(
    (el) => el.getAttribute('style') ?? '',
  )

const pPrOf = (xml: string, text: string) =>
  /<w:pPr>.*?<\/w:pPr>/.exec(xml.split('<w:p>').find((p) => p.includes(`>${text}<`)) ?? '')?.[0]

describe('list geometry for numbering inherited from the paragraph style', () => {
  it('takes indent and hanging from the style, and from the level for a direct w:numPr', async () => {
    const { editor } = await openEditor()
    const [viaStyle, direct] = listStyles(editor)
    expect(viaStyle).toContain('--li-left: 24.95pt')
    expect(viaStyle).toContain('--li-hang: 17.85pt')
    expect(direct).toContain('--li-left: 499.95pt')
    expect(direct).toContain('--li-hang: 18pt')
    editor.destroy()
  })

  it('keeps the paragraph properties of both headings through an edit and save', async () => {
    const { parsed, editor } = await openEditor()
    const end = editor.state.doc.content.size - 1
    editor.commands.insertContentAt(end, ' edited')
    editor.commands.insertContentAt(2, 'x')
    const plan = pmDocToSavePlan(editor.getJSON() as never, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(xml).toContain('Body edited')
    expect(pPrOf(xml, 'Vxia style')).toBe('<w:pPr><w:pStyle w:val="Bar"/></w:pPr>')
    expect(pPrOf(xml, 'Direct')).toBe(`<w:pPr><w:pStyle w:val="Bar"/>${DIRECT_NUMPR}</w:pPr>`)
    editor.destroy()
  })

  it('adds no w:numPr when the style-numbered heading gets a format edit', async () => {
    const { parsed, editor } = await openEditor()
    const heading = editor.state.doc.firstChild!
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, { ...heading.attrs, align: 'center' }),
    )
    const plan = pmDocToSavePlan(editor.getJSON() as never, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(pPrOf(xml, 'Via style')).toBe(
      '<w:pPr><w:pStyle w:val="Bar"/><w:jc w:val="center"/></w:pPr>',
    )
    editor.destroy()
  })
})
