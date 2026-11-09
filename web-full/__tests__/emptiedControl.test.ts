// @vitest-environment jsdom
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import { TextSelection } from '@tiptap/pm/state'
import { EMPTY_CONTROL_TEXT, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../../apps/docs/src/renderer/editor/convert'
import { editorExtensions } from '../../apps/docs/src/renderer/editor/extensions'
import { placeholderAt } from '../placeholderText'

const STYLE =
  '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:rPr><w:color w:val="808080"/></w:rPr></w:style>'
const PLACEHOLDER = 'Vor-/Nachname(n) / Position(en)'
const control = (id: string) =>
  `<w:sdt><w:sdtPr><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:alias w:val="Vertreter"/><w:id w:val="${id}"/>` +
  '<w:placeholder><w:docPart w:val="DefaultPlaceholder_1081868574"/></w:placeholder><w:showingPlcHdr/><w:text/></w:sdtPr>' +
  `<w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t xml:space="preserve">${PLACEHOLDER}</w:t></w:r></w:sdtContent></w:sdt>`
const IN_CELL =
  '<w:tbl><w:tblGrid><w:gridCol w:w="4200"/><w:gridCol w:w="5400"/></w:tblGrid><w:tr>' +
  '<w:tc><w:tcPr><w:tcW w:w="4200" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>gesetzlich vertreten durch</w:t></w:r></w:p></w:tc>' +
  `<w:tc><w:tcPr><w:tcW w:w="5400" w:type="dxa"/></w:tcPr><w:p>${control('103')}</w:p></w:tc></w:tr></w:tbl>`
const AFTER_LABEL = `<w:p><w:r><w:t xml:space="preserve">Vertreten durch: </w:t></w:r>${control('104')}</w:p>`

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((editor) => editor.destroy()))

function open(blocks: Awaited<ReturnType<typeof parseDocx>>['blocks']) {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(blocks) as never)
  editors.push(editor)
  return editor
}

const textPos = (editor: Editor, text: string) => {
  let found = -1
  editor.state.doc.descendants((node, pos) => {
    if (found === -1 && node.isText && node.text?.includes(text))
      found = pos + node.text.indexOf(text)
  })
  return found
}

/** the paragraph end of the paragraph holding `text`, or of the first emptied control's in a table cell */
function paragraphEnd(editor: Editor, text: string | null) {
  let end = -1
  editor.state.doc.descendants((node, pos, parent) => {
    const matches =
      text === null
        ? parent?.type.name.includes('Cell') && node.textContent === EMPTY_CONTROL_TEXT
        : node.textContent.startsWith(text)
    if (end === -1 && node.isTextblock && matches) end = pos + node.nodeSize - 1
    return end === -1
  })
  return end
}

async function saved(parsed: Awaited<ReturnType<typeof parseDocx>>, editor: Editor) {
  return saveDocx(parsed, pmDocToSavePlan(editor.getJSON() as never, parsed.blocks).saveBlocks)
}

const controlTexts = async (bytes: Uint8Array) => {
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
  return [
    ...new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagName('w:sdt'),
  ].map((sdt) => [...sdt.getElementsByTagName('w:t')].map((t) => t.textContent).join(''))
}

describe('typing into a control a previous save emptied', () => {
  it.each([
    ['in a table cell', IN_CELL, null],
    ['after a label in its paragraph', AFTER_LABEL, 'Vertreten durch: '],
  ])('%s puts the text inside the control', async (_, body, label) => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: body, extraStylesXml: STYLE }))
    const emptying = open(parsed.blocks)
    const range = placeholderAt(emptying.state.doc, textPos(emptying, PLACEHOLDER))!
    emptying.view.dispatch(emptying.state.tr.delete(range.from, range.to))
    const emptied = await saved(parsed, emptying)
    expect(await controlTexts(emptied)).toEqual([''])

    const reparsed = await parseDocx(emptied)
    const typing = open(reparsed.blocks)
    const at = paragraphEnd(typing, label)
    typing.view.dispatch(typing.state.tr.setSelection(TextSelection.create(typing.state.doc, at)))
    typing.view.dispatch(typing.state.tr.insertText('Erika Muster, Geschäftsführerin'))

    expect(await controlTexts(await saved(reparsed, typing))).toEqual([
      'Erika Muster, Geschäftsführerin',
    ])
  })

  it('a control an earlier save left without any w:t takes the text inside it', async () => {
    const bare = IN_CELL.replace(`<w:t xml:space="preserve">${PLACEHOLDER}</w:t>`, '').replace(
      '<w:showingPlcHdr/>',
      '',
    )
    const parsed = await parseDocx(await buildDocx({ bodyXml: bare, extraStylesXml: STYLE }))
    expect(await controlTexts(await buildDocx({ bodyXml: bare, extraStylesXml: STYLE }))).toEqual([
      '',
    ])
    const typing = open(parsed.blocks)
    const at = paragraphEnd(typing, null)
    typing.view.dispatch(typing.state.tr.setSelection(TextSelection.create(typing.state.doc, at)))
    typing.view.dispatch(typing.state.tr.insertText('Erika Muster'))

    expect(await controlTexts(await saved(parsed, typing))).toEqual(['Erika Muster'])
  })
})
