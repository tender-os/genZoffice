/** A blank legacy text form field (its result is en spaces, Word's empty form field) opens as text carrying
 * the field's mark (Vound fork). An untouched save leaves the document byte-identical; a typed result lands
 * inside the field result with the field machinery kept. */
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const BLANK_FORMTEXT =
  '<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Text2"/><w:enabled/><w:textInput/></w:ffData></w:fldChar></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>\u2002</w:t></w:r><w:r><w:t>\u2002</w:t></w:r><w:r><w:t>\u2002</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
const BODY =
  '<w:p><w:r><w:t>Firma:</w:t></w:r></w:p>' +
  BLANK_FORMTEXT +
  '<w:p><w:r><w:t>Ort:</w:t></w:r></w:p>'
const FIELD_MACHINERY = [
  /<w:ffData>/g,
  /FORMTEXT/g,
  /fldCharType="begin"/g,
  /fldCharType="separate"/g,
  /fldCharType="end"/g,
]

const documentXmlOf = async (bytes: Uint8Array) =>
  (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')

async function open() {
  const bytes = await buildDocx({ bodyXml: BODY })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { bytes, parsed, editor }
}

/** the text between the field's separate and end */
const fieldResult = (xml: string) => {
  const separate = 'fldCharType="separate"/>'
  const from = xml.indexOf(separate) + separate.length
  return xml
    .slice(from, xml.indexOf('<w:fldChar w:fldCharType="end"', from))
    .replace(/<[^>]+>/g, '')
}

describe('a blank legacy text form field through the editor', () => {
  it('opens as text and saves untouched byte for byte', async () => {
    const { bytes, parsed, editor } = await open()
    expect(editor.state.doc.child(1).type.name).toBe('docParagraph')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    editor.destroy()
    expect(plan.changedCount).toBe(0)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    expect(await documentXmlOf(saved)).toBe(await documentXmlOf(bytes))
  })

  it('saves a typed result inside the field result and keeps the field machinery', async () => {
    const { bytes, parsed, editor } = await open()
    const start = editor.state.doc.child(0).nodeSize + 1
    editor.view.dispatch(editor.state.tr.insertText('Muster GmbH', start, start + 3))
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    editor.destroy()
    expect(plan.changedCount).toBe(1)
    const savedXml = await documentXmlOf(await saveDocx(parsed, plan.saveBlocks))
    const sourceXml = await documentXmlOf(bytes)
    expect(fieldResult(savedXml)).toBe('Muster GmbH')
    expect(savedXml).toContain('<w:t>Firma:</w:t>')
    expect(savedXml).toContain('<w:t>Ort:</w:t>')
    for (const pattern of FIELD_MACHINERY)
      expect(savedXml.match(pattern)?.length).toBe(sourceXml.match(pattern)?.length)
  })

  it("saves a value typed into a line of several blank fields into the first field's result", async () => {
    const fieldRuns = BLANK_FORMTEXT.slice('<w:p>'.length, -'</w:p>'.length)
    const bytes = await buildDocx({
      bodyXml: `<w:p><w:r><w:t>Frau/Herr</w:t></w:r></w:p><w:p>${fieldRuns.repeat(3)}</w:p>`,
    })
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const start = editor.state.doc.child(0).nodeSize + 1
    editor.view.dispatch(editor.state.tr.insertText('Erika Mustermann', start, start + 3))
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    editor.destroy()
    const savedXml = await documentXmlOf(await saveDocx(parsed, plan.saveBlocks))
    const results = savedXml
      .split('fldCharType="separate"/>')
      .slice(1)
      .map((part) =>
        part.slice(0, part.indexOf('<w:fldChar w:fldCharType="end"')).replace(/<[^>]+>/g, ''),
      )
    expect(results).toEqual(['Erika Mustermann', '\u2002'.repeat(3), '\u2002'.repeat(3)])
  })
})
