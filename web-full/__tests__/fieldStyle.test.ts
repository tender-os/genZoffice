// @vitest-environment jsdom
import { Editor } from '@tiptap/core'
import { undo } from '@tiptap/pm/history'
import { TextSelection } from '@tiptap/pm/state'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type ParsedDocFull } from '@genoffice/docx-engine'
import { buildDocx } from '../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../../apps/docs/src/renderer/editor/convert'
import { editorExtensions } from '../../apps/docs/src/renderer/editor/extensions'
import type { BlockRef } from '../embed-protocol'
import { applyFieldStyle } from '../fieldStyle'
import { HOST_WRITE_META } from '../fieldWrite'
import { formTextPlugin } from '../formTextField'

const run = (text: string, rPr = '') => `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`
const CALIBRI_11 = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
const cell = (inner: string) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p>${inner}</w:p></w:tc>`
const fieldOpen =
  '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="F"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
const FIELD_END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
const BODY =
  `<w:p>${run('Firma: ')}${run('Muster GmbH', CALIBRI_11)}</w:p>` +
  `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
  `<w:tr>${cell(run('Ort'))}${cell(`${fieldOpen}${run('Berlin')}${FIELD_END}`)}</w:tr></w:tbl>` +
  `<w:p>${run('Ende')}</w:p>`
const firm: BlockRef = { paragraph: 0, scope: 'paragraph' }
const city: BlockRef = { paragraph: 2, scope: 'cell' }

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((editor) => editor.destroy()))

async function open(bodyXml = BODY) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const element = document.createElement('div')
  document.body.append(element)
  const editor = new Editor({ element, extensions: editorExtensions })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  editor.registerPlugin(formTextPlugin, (plugin, plugins) => [plugin, ...plugins])
  editors.push(editor)
  return { parsed, editor }
}

async function reopened(parsed: ParsedDocFull, editor: Editor) {
  const plan = pmDocToSavePlan(editor.getJSON() as never, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
  return { xml, again: await open(xml.slice(xml.indexOf('<w:body>') + 8, xml.indexOf('<w:sectPr'))) }
}

/** types the way the reviewer does: through the editor's text input handlers, else as plain input */
function type(editor: Editor, text: string) {
  const { from, to } = editor.state.selection
  const handled = editor.view.someProp('handleTextInput', (handle) => handle(editor.view, from, to, text, () => editor.state.tr.insertText(text, from, to)))
  if (!handled) editor.view.dispatch(editor.state.tr.insertText(text, from, to))
}

const styleOfText = (editor: Editor, text: string) => {
  let attrs: Record<string, unknown> | undefined
  editor.state.doc.descendants((node) => {
    if (node.isText && node.text === text) attrs = node.marks.find((mark) => mark.type.name === 'docTextStyle')?.attrs
  })
  return attrs
}
const faceOfText = (editor: Editor, text: string) => styleOfText(editor, text)?.fontAscii

async function savedXml(parsed: ParsedDocFull, editor: Editor) {
  const plan = pmDocToSavePlan(editor.getJSON() as never, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
}

const runOf = (xml: string, text: string) => {
  const end = xml.indexOf(`>${text}</w:t>`)
  return xml.slice(xml.lastIndexOf('<w:r>', end), end)
}

const selectText = (editor: Editor, text: string) => {
  let from = -1
  editor.state.doc.descendants((node, pos) => {
    if (from === -1 && node.isText && node.text!.includes(text)) from = pos + node.text!.indexOf(text)
  })
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, from + text.length)))
}

describe('applyFieldStyle', () => {
  it('restyles the whole field when nothing in it is selected, and saves the run properties', async () => {
    const { parsed, editor } = await open()
    const result = applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: { fontFamily: 'Arial', fontSize: 12 } })
    expect(result).toEqual({ ok: true, style: { fontFamily: 'Arial', fontSize: 12 } })
    const xml = await savedXml(parsed, editor)
    expect(runOf(xml, 'Muster GmbH')).toContain('w:ascii="Arial"')
    expect(runOf(xml, 'Muster GmbH')).toContain('w:hAnsi="Arial"')
    expect(runOf(xml, 'Muster GmbH')).toContain('<w:sz w:val="24"/>')
    expect(runOf(xml, 'Muster GmbH')).toContain('<w:szCs w:val="24"/>')
    expect(runOf(xml, 'Firma: ')).not.toContain('Arial')
  })

  it('restyles only the selection when it lies inside the field', async () => {
    const { parsed, editor } = await open()
    selectText(editor, 'GmbH')
    applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: { fontSize: 14 } })
    const xml = await savedXml(parsed, editor)
    expect(runOf(xml, 'GmbH')).toContain('<w:sz w:val="28"/>')
    expect(runOf(xml, 'Muster ')).toContain('<w:sz w:val="22"/>')
  })

  it('keeps what the run had and changes only the property asked for', async () => {
    const { parsed, editor } = await open()
    applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: { fontSize: 9 } })
    const field = runOf(await savedXml(parsed, editor), 'Muster GmbH')
    expect(field).toContain('w:ascii="Calibri"')
    expect(field).toContain('<w:sz w:val="18"/>')
  })

  it('restyles a form text field in a table cell, found by its gap', async () => {
    const { parsed, editor } = await open()
    const result = applyFieldStyle(editor.view, { block: city, replaces: 'gap', gap: { before: '', after: '' }, style: { fontFamily: 'Courier New' } })
    expect(result.ok).toBe(true)
    const xml = await savedXml(parsed, editor)
    expect(runOf(xml, 'Berlin')).toContain('w:ascii="Courier New"')
    expect(xml).toContain('FORMTEXT')
  })

  it('is one undoable host write', async () => {
    const { editor } = await open()
    const metas: unknown[] = []
    editor.on('transaction', ({ transaction }) => { if (transaction.docChanged) metas.push(transaction.getMeta(HOST_WRITE_META)) })
    const before = editor.state.doc
    applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: { fontFamily: 'Arial', fontSize: 12 } })
    expect(metas).toEqual([true])
    undo(editor.state, editor.view.dispatch)
    expect(editor.state.doc.eq(before)).toBe(true)
  })

  it('answers with the style the field shows and changes nothing for an empty style', async () => {
    const { editor } = await open()
    const before = editor.state.doc
    expect(applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: {} })).toEqual({ ok: true, style: { fontFamily: 'Calibri', fontSize: 11 } })
    expect(editor.state.doc).toBe(before)
  })

  it('leaves out a property the field shows in more than one way', async () => {
    const { editor } = await open()
    selectText(editor, 'GmbH')
    applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: { fontSize: 14 } })
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1)))
    expect(applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: {} })).toEqual({ ok: true, style: { fontFamily: 'Calibri' } })
  })

  it('restyles a content control in a paragraph and in a cell, through save and reload', async () => {
    const control = (text: string) => `<w:sdt><w:sdtPr><w:id w:val="7"/><w:text/></w:sdtPr><w:sdtContent>${run(text, CALIBRI_11)}</w:sdtContent></w:sdt>`
    const body =
      `<w:p>${run('Ort: ')}${control('Berlin')}</w:p>` +
      `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
      `<w:tr>${cell(run('Firma'))}${cell(control('Muster GmbH'))}</w:tr></w:tbl><w:p>${run('Ende')}</w:p>`
    const { parsed, editor } = await open(body)
    expect(applyFieldStyle(editor.view, { block: firm, previous: 'Berlin', style: { fontFamily: 'Arial' } }).ok).toBe(true)
    expect(applyFieldStyle(editor.view, { block: { paragraph: 2, scope: 'cell' }, previous: 'Muster GmbH', style: { fontSize: 9 } }).ok).toBe(true)
    const { xml, again } = await reopened(parsed, editor)
    expect(xml.match(/<w:sdt>/g)).toHaveLength(2)
    expect(runOf(xml, 'Berlin')).toContain('w:ascii="Arial"')
    expect(runOf(xml, 'Muster GmbH')).toContain('<w:sz w:val="18"/>')
    expect(applyFieldStyle(again.editor.view, { block: firm, previous: 'Berlin', style: {} })).toEqual({ ok: true, style: { fontFamily: 'Arial', fontSize: 11 } })
    expect(applyFieldStyle(again.editor.view, { block: { paragraph: 2, scope: 'cell' }, previous: 'Muster GmbH', style: {} })).toEqual({ ok: true, style: { fontFamily: 'Calibri', fontSize: 9 } })
  })

  it('restyles the result of a form field shown as one block, and saves it there only', async () => {
    const dropDown = '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="D"/><w:ddList><w:listEntry w:val="ja"/></w:ddList></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' + run('ja') + FIELD_END
    const { parsed, editor } = await open(`<w:p>${run('Firma: ', CALIBRI_11)}${fieldOpen}${run('Muster GmbH', CALIBRI_11)}${FIELD_END}${run(' ')}${dropDown}</w:p>`)
    expect(editor.state.doc.child(0).attrs.fieldDisplay).not.toBeNull()
    const result = applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: { fontFamily: 'Arial', fontSize: 12 } })
    expect(result).toEqual({ ok: true, style: { fontFamily: 'Arial', fontSize: 12 } })
    const shown = editor.view.nodeDOM(0) as HTMLElement
    expect(shown.textContent).toContain('Muster GmbH')
    const { xml, again } = await reopened(parsed, editor)
    expect(runOf(xml, 'Muster GmbH')).toContain('w:ascii="Arial"')
    expect(runOf(xml, 'Muster GmbH')).toContain('<w:sz w:val="24"/>')
    expect(runOf(xml, 'Firma: ')).toContain('w:ascii="Calibri"')
    expect(xml).toContain('FORMDROPDOWN')
    expect(applyFieldStyle(again.editor.view, { block: firm, previous: 'Muster GmbH', style: {} })).toEqual({ ok: true, style: { fontFamily: 'Arial', fontSize: 12 } })
  })

  it('restyles an empty form field the field index names, and the text typed into it takes the style', async () => {
    const blank = `${fieldOpen}${run('\u2002'.repeat(5))}${FIELD_END}`
    const { parsed, editor } = await open(`<w:p>${run('Ort: ')}${blank}${run(' Datum: ')}${blank}</w:p>`)
    expect(applyFieldStyle(editor.view, { block: firm, previous: '\u2002'.repeat(5), formText: 1, style: { fontFamily: 'Arial' } }).ok).toBe(true)
    let end = 0
    editor.state.doc.descendants((node, pos) => { if (node.isText) end = pos + node.nodeSize })
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, end)))
    type(editor, 'Mai')
    expect(faceOfText(editor, 'Mai')).toBe('Arial')
    const { xml } = await reopened(parsed, editor)
    expect(runOf(xml, 'Mai')).toContain('w:ascii="Arial"')
  })

  it.each([
    ['its index', { formText: 0 }],
    ['being the one empty field of its line', {}],
  ])('restyles a form field saved empty, found by %s, and the text typed into it takes the style', async (_, hint) => {
    const { parsed, editor } = await open(`<w:p>${run('Ort: ')}${fieldOpen}${run('')}${FIELD_END}</w:p>`)
    expect(applyFieldStyle(editor.view, { block: firm, previous: '', ...hint, style: { fontSize: 14 } }).ok).toBe(true)
    type(editor, 'Köln')
    expect(styleOfText(editor, 'Köln')?.sizeHalfPoints).toBe(28)
    const { xml } = await reopened(parsed, editor)
    expect(runOf(xml, 'Köln')).toContain('<w:sz w:val="28"/>')
    expect(xml).toContain('FORMTEXT')
  })

  it('stores the style at the end of a label line whose field has no text yet, for the text typed next', async () => {
    const { parsed, editor } = await open(`<w:p>${run('Name: ')}</w:p>`)
    const result = applyFieldStyle(editor.view, { block: firm, previous: '', style: { fontFamily: 'Courier New' } })
    expect(result.ok).toBe(true)
    type(editor, 'Muster')
    expect(faceOfText(editor, 'Muster')).toBe('Courier New')
    const { xml } = await reopened(parsed, editor)
    expect(runOf(xml, 'Muster')).toContain('w:ascii="Courier New"')
  })

  it('refuses a target that names no place, and leaves the document as it was', async () => {
    const { editor } = await open()
    const before = editor.state.doc
    expect(applyFieldStyle(editor.view, { block: firm, previous: 'Beispiel AG', style: { fontSize: 12 } })).toEqual({ ok: false, reason: 'absent' })
    expect(applyFieldStyle(editor.view, { block: { paragraph: 99, scope: 'paragraph' }, previous: 'x', style: { fontSize: 12 } })).toEqual({ ok: false, reason: 'noBlock' })
    expect(editor.state.doc).toBe(before)
  })
})
