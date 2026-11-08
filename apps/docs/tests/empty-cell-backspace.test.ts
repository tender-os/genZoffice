/**
 * Backspace in an emptied table cell paragraph (Vound fork): Word leaves the paragraph as it is, its
 * properties included; only its runs went with the text. Tiptap's own Backspace turned it into a default
 * paragraph, so the cell lost its alignment, its style and its source paragraph index.
 */
import { Editor } from '@tiptap/core'
import { CellSelection } from '@tiptap/pm/tables'
import { TextSelection } from '@tiptap/pm/state'
import { afterEach, describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const STYLE = '<w:style w:type="paragraph" w:styleId="Formular"><w:name w:val="Formular"/></w:style>'
const TABLE =
  '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>' +
  '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Firma</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="Formular"/><w:jc w:val="center"/></w:pPr><w:r><w:t>Muster</w:t></w:r></w:p></w:tc>' +
  '</w:tr></w:tbl><w:p/>'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

async function open() {
  const parsed = await parseDocx(await buildDocx({ bodyXml: TABLE, extraStylesXml: STYLE }))
  const editor = new Editor({ element: document.createElement('div'), extensions: editorExtensions })
  editors.push(editor)
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  return editor
}

/** where the value cell's paragraph starts */
function valuePos(editor: Editor) {
  const cells: number[] = []
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name === 'docTableCell') cells.push(pos)
  })
  return cells[1] + 1
}

const valueAttrs = (editor: Editor) => {
  const node = editor.state.doc.nodeAt(valuePos(editor))!
  return { text: node.textContent, styleId: node.attrs.styleId, align: node.attrs.align, docxPara: node.attrs.docxPara }
}

describe('Backspace in an emptied table cell paragraph', () => {
  it('keeps the paragraph and its properties', async () => {
    const editor = await open()
    const before = valueAttrs(editor)
    expect(before).toEqual({ text: 'Muster', styleId: 'Formular', align: 'center', docxPara: 1 })
    const pos = valuePos(editor)
    editor.view.dispatch(editor.state.tr.delete(pos + 1, pos + 7))
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos + 1)))
    editor.commands.keyboardShortcut('Backspace')
    expect(valueAttrs(editor)).toEqual({ ...before, text: '' })
  })

  it('keeps them on Delete in the emptied paragraph', async () => {
    const editor = await open()
    const before = valueAttrs(editor)
    const pos = valuePos(editor)
    editor.view.dispatch(editor.state.tr.delete(pos + 1, pos + 7))
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos + 1)))
    editor.commands.keyboardShortcut('Delete')
    expect(valueAttrs(editor)).toEqual({ ...before, text: '' })
  })

  it('keeps them when a cell selection is cleared', async () => {
    const editor = await open()
    const before = valueAttrs(editor)
    const cell = editor.state.doc.resolve(valuePos(editor)).before()
    editor.view.dispatch(editor.state.tr.setSelection(CellSelection.create(editor.state.doc, cell)))
    editor.commands.keyboardShortcut('Backspace')
    expect(valueAttrs(editor)).toMatchObject({ text: '', styleId: before.styleId, align: before.align })
  })

  it('keeps them and takes the typed text when typing replaces a cell selection', async () => {
    const editor = await open()
    const before = valueAttrs(editor)
    const cell = editor.state.doc.resolve(valuePos(editor)).before()
    editor.view.dispatch(editor.state.tr.setSelection(CellSelection.create(editor.state.doc, cell)))
    const { view } = editor
    const { from, to } = view.state.selection
    // what ProseMirror does with a typed character: the handlers first, else its own insertText
    const fallback = () => view.state.tr.insertText('Neu')
    if (!view.someProp('handleTextInput', (handle) => handle(view, from, to, 'Neu', fallback))) view.dispatch(fallback())
    expect(valueAttrs(editor)).toMatchObject({ text: 'Neu', styleId: before.styleId, align: before.align })
  })
})
