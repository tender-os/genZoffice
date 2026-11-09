/**
 * A split source paragraph (Vound fork): Enter inside a paragraph keeps its source index (docxPara) on the
 * first half and marks the second as its continuation (continuesPara), so the index stays unambiguous while
 * the field's marker can cover both halves; joining them again gives the one paragraph back. The save keeps
 * both halves where the paragraph was.
 */
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

const editorWith = (paragraphs: Array<[string, number | null]>) => {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: paragraphs.map(([text, docxPara]) => ({ type: 'docParagraph', attrs: { docxPara }, content: text ? [{ type: 'text', text }] : [] })),
    } as never,
  })
  editors.push(editor)
  return editor
}

const indices = (editor: Editor) => {
  const out: Array<[string, unknown, unknown]> = []
  editor.state.doc.descendants((node) => {
    if (node.isTextblock) out.push([node.textContent, node.attrs.docxPara, node.attrs.continuesPara])
    return !node.isTextblock
  })
  return out
}

/** where `text` starts */
const at = (editor: Editor, text: string) => {
  let found = -1
  editor.state.doc.descendants((node, pos) => {
    if (node.isText && found === -1 && node.text!.includes(text)) found = pos + node.text!.indexOf(text)
  })
  return found
}

describe('a split source paragraph', () => {
  it('keeps the index on the first half and marks the second as its continuation', () => {
    const editor = editorWith([['Vorher', 1], ['Muster GmbH', 2], ['Nachher', 3]])
    editor.commands.setTextSelection(at(editor, 'GmbH'))
    editor.commands.keyboardShortcut('Enter')
    expect(indices(editor)).toEqual([['Vorher', 1, null], ['Muster ', 2, null], ['GmbH', null, 2], ['Nachher', 3, null]])
    editor.commands.insertContentAt(at(editor, 'GmbH'), 'Beispiel ')
    editor.commands.setTextSelection(at(editor, 'Beispiel') + 'Beispiel GmbH'.length)
    editor.commands.keyboardShortcut('Enter')
    expect(indices(editor).map(([, docxPara, continuesPara]) => [docxPara, continuesPara])).toEqual([[1, null], [2, null], [null, 2], [null, 2], [3, null]])
  })

  it('gives the one paragraph back when the halves are joined again', () => {
    const editor = editorWith([['Muster GmbH', 2]])
    editor.commands.setTextSelection(at(editor, 'GmbH'))
    editor.commands.keyboardShortcut('Enter')
    editor.commands.keyboardShortcut('Backspace')
    expect(indices(editor)).toEqual([['Muster GmbH', 2, null]])
  })

  it('moves the index to the continuation when the first half goes', () => {
    const editor = editorWith([['Vorher', 1], ['Muster GmbH', 2]])
    editor.commands.setTextSelection(at(editor, 'GmbH'))
    editor.commands.keyboardShortcut('Enter')
    editor.commands.deleteRange({ from: at(editor, 'Vorher') + 'Vorher'.length, to: at(editor, 'GmbH') - 1 })
    expect(indices(editor)).toEqual([['Vorher', 1, null], ['GmbH', 2, null]])
  })

  it('gives no index to a copy of a paragraph elsewhere', () => {
    const editor = editorWith([['Muster GmbH', 2], ['Nachher', 3]])
    const copy = editor.schema.nodes.docParagraph.create({ docxPara: 2 }, editor.schema.text('Kopie'))
    editor.view.dispatch(editor.state.tr.insert(editor.state.doc.content.size, copy))
    expect(indices(editor)).toEqual([['Muster GmbH', 2, null], ['Nachher', 3, null], ['Kopie', null, null]])
  })

  it('saves both halves of a split cell paragraph inside its content control, and they reparse', async () => {
    const sdt = (inside: string) =>
      `<w:sdt><w:sdtPr><w:tag w:val="form:3"/><w:id w:val="3"/><w:richText/></w:sdtPr><w:sdtContent>${inside}</w:sdtContent></w:sdt>`
    const table =
      '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
      `${sdt('<w:p><w:r><w:t>Muster GmbH</w:t></w:r></w:p>')}</w:tc></w:tr></w:tbl><w:p/>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: table }))
    const editor = new Editor({ element: document.createElement('div'), extensions: editorExtensions })
    editors.push(editor)
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    editor.commands.setTextSelection(at(editor, 'GmbH'))
    editor.commands.keyboardShortcut('Enter')
    editor.commands.insertContentAt(at(editor, 'Muster') + 'Muster'.length, ' Bau')
    editor.commands.insertContentAt(at(editor, 'GmbH') + 'GmbH'.length, ' & Co. KG')
    expect(indices(editor).slice(0, 2).map(([text, , continuesPara]) => [text, continuesPara])).toEqual([['Muster Bau ', null], ['GmbH & Co. KG', 0]])
    const out = await saveDocx(parsed, pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks)
    const xml = await (await JSZip.loadAsync(out)).file('word/document.xml')!.async('string')
    expect(xml).toContain(sdt('<w:p><w:r><w:t xml:space="preserve">Muster Bau </w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">GmbH &amp; Co. KG</w:t></w:r></w:p>'))
    const reparsed = await parseDocx(out)
    expect(reparsed.blocks[0].table!.rows[0][0].richParas!.map((p) => p.runs.map((run) => run.text).join(''))).toEqual(['Muster Bau ', 'GmbH & Co. KG'])
  })
})
