// Shared fixtures for the field style tests: a real editor on a built document, its save and reopen.
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import JSZip from 'jszip'
import { afterEach } from 'vitest'
import { parseDocx, saveDocx, type ParsedDocFull } from '@genoffice/docx-engine'
import { buildDocx } from '../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../../apps/docs/src/renderer/editor/convert'
import { editorExtensions } from '../../apps/docs/src/renderer/editor/extensions'
import type { BlockRef } from '../embed-protocol'
import { formTextPlugin } from '../formTextField'

export const run = (text: string, rPr = '') => `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`
export const CALIBRI_11 = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
export const cell = (inner: string) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p>${inner}</w:p></w:tc>`
export const fieldOpen =
  '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="F"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
export const FIELD_END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
export const BODY =
  `<w:p>${run('Firma: ')}${run('Muster GmbH', CALIBRI_11)}</w:p>` +
  `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
  `<w:tr>${cell(run('Ort'))}${cell(`${fieldOpen}${run('Berlin')}${FIELD_END}`)}</w:tr></w:tbl>` +
  `<w:p>${run('Ende')}</w:p>`
export const firm: BlockRef = { paragraph: 0, scope: 'paragraph' }
export const city: BlockRef = { paragraph: 2, scope: 'cell' }

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((editor) => editor.destroy()))

export async function open(bodyXml = BODY) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const element = document.createElement('div')
  document.body.append(element)
  const editor = new Editor({ element, extensions: editorExtensions })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  editor.registerPlugin(formTextPlugin, (plugin, plugins) => [plugin, ...plugins])
  editors.push(editor)
  return { parsed, editor }
}

export async function reopened(parsed: ParsedDocFull, editor: Editor) {
  const plan = pmDocToSavePlan(editor.getJSON() as never, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
  return { xml, again: await open(xml.slice(xml.indexOf('<w:body>') + 8, xml.indexOf('<w:sectPr'))) }
}

/** types the way the reviewer does: through the editor's text input handlers, else as plain input */
export function type(editor: Editor, text: string) {
  const { from, to } = editor.state.selection
  const handled = editor.view.someProp('handleTextInput', (handle) => handle(editor.view, from, to, text, () => editor.state.tr.insertText(text, from, to)))
  if (!handled) editor.view.dispatch(editor.state.tr.insertText(text, from, to))
}

export const styleOfText = (editor: Editor, text: string) => {
  let attrs: Record<string, unknown> | undefined
  editor.state.doc.descendants((node) => {
    if (node.isText && node.text === text) attrs = node.marks.find((mark) => mark.type.name === 'docTextStyle')?.attrs
  })
  return attrs
}
export const faceOfText = (editor: Editor, text: string) => styleOfText(editor, text)?.fontAscii

export async function savedXml(parsed: ParsedDocFull, editor: Editor) {
  const plan = pmDocToSavePlan(editor.getJSON() as never, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
}

export const runOf = (xml: string, text: string) => {
  const end = xml.indexOf(`>${text}</w:t>`)
  return xml.slice(xml.lastIndexOf('<w:r>', end), end)
}

export const selectText = (editor: Editor, text: string) => {
  let from = -1
  editor.state.doc.descendants((node, pos) => {
    if (from === -1 && node.isText && node.text!.includes(text)) from = pos + node.text!.indexOf(text)
  })
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, from + text.length)))
}

