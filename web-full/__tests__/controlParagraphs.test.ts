// @vitest-environment jsdom
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx, type ParsedDocFull } from '@genoffice/docx-engine'
import { buildDocx } from '../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../../apps/docs/src/renderer/editor/convert'
import { editorExtensions } from '../../apps/docs/src/renderer/editor/extensions'
import { editScopePlugin } from '../editScope'
import { placeholderAt } from '../placeholderText'

const STYLE =
  '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:rPr><w:color w:val="808080"/></w:rPr></w:style>'
const control = (id: string, text: string) =>
  `<w:p><w:sdt><w:sdtPr><w:id w:val="${id}"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>${text}</w:t></w:r></w:sdtContent></w:sdt></w:p>`
const CONTROLS = `${control('1', 'Datum eingeben')}${control('2', 'Ort eintragen')}`
const blockControl = (id: string, text: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="${id}"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:sdtContent></w:sdt>`
const BLOCK_CONTROLS = `${blockControl('3', 'Musterstadt')}${blockControl('4', 'Musterweg 1')}`
const IN_CELL = `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${CONTROLS}</w:tc></w:tr></w:tbl>`

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((editor) => editor.destroy()))

async function open(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, extraStylesXml: STYLE }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  editor.registerPlugin(editScopePlugin(vi.fn()))
  editors.push(editor)
  return { parsed, editor }
}

const textPos = (editor: Editor, text: string) => {
  let found = -1
  editor.state.doc.descendants((node, pos) => {
    if (found === -1 && node.isText && node.text?.includes(text))
      found = pos + node.text.indexOf(text)
  })
  return found
}

function selectPlaceholder(editor: Editor, text: string) {
  const range = placeholderAt(editor.state.doc, textPos(editor, text))!
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, range.from, range.to)),
  )
}

async function savedControls(parsed: ParsedDocFull, editor: Editor) {
  const saved = await saveDocx(
    parsed,
    pmDocToSavePlan(editor.getJSON() as never, parsed.blocks).saveBlocks,
  )
  const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
  return new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagName('w:sdt')
    .length
}

const controlParagraphTexts = (editor: Editor) => {
  const texts: string[] = []
  editor.state.doc.descendants((node) => {
    if (node.attrs.holdsControl === true || node.attrs.sdtShell) texts.push(node.textContent)
    return !node.isTextblock
  })
  return texts
}

describe('content controls through a review edit', () => {
  it.each([
    ['a table cell', IN_CELL],
    ['the body', CONTROLS],
  ])(
    'in %s, Backspace at the start of an emptied control keeps its paragraph and both controls',
    async (_, body) => {
      const { parsed, editor } = await open(body)
      selectPlaceholder(editor, 'Datum eingeben')
      editor.view.dispatch(editor.state.tr.insertText('09.11.2026'))
      selectPlaceholder(editor, 'Ort eintragen')
      editor.commands.keyboardShortcut('Backspace')
      editor.commands.keyboardShortcut('Backspace')

      expect(controlParagraphTexts(editor)).toEqual(['09.11.2026', ''])
      expect(await savedControls(parsed, editor)).toBe(2)
    },
  )

  it('Enter in a control paragraph does nothing', async () => {
    const { parsed, editor } = await open(IN_CELL)
    selectPlaceholder(editor, 'Datum eingeben')
    editor.view.dispatch(editor.state.tr.insertText('09.11.2026'))
    editor.commands.enter()

    expect(controlParagraphTexts(editor)).toEqual(['09.11.2026', 'Ort eintragen'])
    expect(await savedControls(parsed, editor)).toBe(2)
  })

  it('typing over a placeholder still works', async () => {
    const { parsed, editor } = await open(IN_CELL)
    selectPlaceholder(editor, 'Ort eintragen')
    editor.view.dispatch(editor.state.tr.insertText('Berlin'))

    expect(controlParagraphTexts(editor)).toEqual(['Datum eingeben', 'Berlin'])
    expect(await savedControls(parsed, editor)).toBe(2)
  })

  it('Enter inside a block-level control does not split it', async () => {
    const { parsed, editor } = await open(BLOCK_CONTROLS)
    editor.commands.setTextSelection(textPos(editor, 'Musterstadt') + 6)
    editor.commands.enter()

    expect(await savedControls(parsed, editor)).toBe(2)
    expect(controlParagraphTexts(editor)).toEqual(['Musterstadt', 'Musterweg 1'])
  })

  it('Backspace at the start of a block-level control does not join it', async () => {
    const { parsed, editor } = await open(BLOCK_CONTROLS)
    editor.commands.setTextSelection(textPos(editor, 'Musterweg 1'))
    editor.commands.keyboardShortcut('Backspace')

    expect(await savedControls(parsed, editor)).toBe(2)
    expect(controlParagraphTexts(editor)).toEqual(['Musterstadt', 'Musterweg 1'])
  })

  it('typing inside a block-level control keeps it', async () => {
    const { parsed, editor } = await open(BLOCK_CONTROLS)
    editor.commands.setTextSelection(textPos(editor, 'Musterweg 1') + 'Musterweg '.length)
    editor.commands.insertContent('2')

    expect(controlParagraphTexts(editor)).toEqual(['Musterstadt', 'Musterweg 21'])
    expect(await savedControls(parsed, editor)).toBe(2)
  })
})
