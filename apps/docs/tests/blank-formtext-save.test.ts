/** A blank legacy text form field (its result is en spaces, Word's empty form field) opens as an empty
 * text field (Vound fork, parse-fields.ts blankFormTextField). An untouched save leaves the document
 * byte-identical; a typed result lands inside the field result with the field machinery kept. */
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
  it('opens as an empty text field and saves untouched byte for byte', async () => {
    const { bytes, parsed, editor } = await open()
    const field = editor.getJSON().content?.find((node) => node.type === 'docProtected')
    expect(field?.attrs?.fieldDisplay).toMatchObject({ kind: 'text', left: '', emptyLine: true })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    editor.destroy()
    expect(plan.changedCount).toBe(0)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    expect(await documentXmlOf(saved)).toBe(await documentXmlOf(bytes))
  })

  it('saves a typed result inside the field result and keeps the field machinery', async () => {
    const { bytes, parsed, editor } = await open()
    let fieldPos = -1
    editor.state.doc.forEach((node, pos) => {
      if (node.type.name === 'docProtected') fieldPos = pos
    })
    const node = editor.state.doc.nodeAt(fieldPos)!
    // what the editor's field commit writes once the reviewer typed (extensions.ts, wireFieldEditing)
    const { emptyLine: _placeholderLine, ...display } = node.attrs.fieldDisplay
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(fieldPos, undefined, {
        ...node.attrs,
        fieldDisplay: { ...display, left: 'Muster GmbH' },
      }),
    )
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

  it("typing after the placeholder of an empty field replaces it (the editor's own field commit)", async () => {
    const { parsed, editor } = await open()
    /** type `typed` into the field's text the way the reviewer does, then let the editor commit it */
    const typeIntoField = (typed: (shown: string) => string) => {
      const text = editor.view.dom.querySelector<HTMLElement>('.doc-field-text')!
      text.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, button: 0 }))
      text.textContent = typed(text.textContent ?? '')
      text.dispatchEvent(new Event('input', { bubbles: true }))
      window.dispatchEvent(new Event('ai-docs-commit-tables'))
    }
    const display = () => editor.state.doc.child(1).attrs.fieldDisplay
    // an empty line's placeholder NBSP, caret after it
    typeIntoField((shown) => `${shown}Muster`)
    expect(display().left).toBe('Muster')
    expect(display().emptyLine).toBeUndefined()

    // Word's empty result written by the host (five en spaces), caret after it
    const enSpaces = '\u2002'.repeat(5)
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(editor.state.doc.child(0).nodeSize, undefined, {
        ...editor.state.doc.child(1).attrs,
        fieldDisplay: { ...display(), left: enSpaces },
      }),
    )
    typeIntoField((shown) => `${shown}Beispiel AG`)
    expect(display().left).toBe('Beispiel AG')

    // with the caret in the middle of the blank, all of it goes
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(editor.state.doc.child(0).nodeSize, undefined, {
        ...editor.state.doc.child(1).attrs,
        fieldDisplay: { ...display(), left: enSpaces },
      }),
    )
    typeIntoField((shown) => `${shown.slice(0, 2)}Mitte${shown.slice(2)}`)
    expect(display().left).toBe('Mitte')
    typeIntoField(() => 'Beispiel AG')
    expect(display().left).toBe('Beispiel AG')

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    editor.destroy()
    expect(fieldResult(await documentXmlOf(await saveDocx(parsed, plan.saveBlocks)))).toBe(
      'Beispiel AG',
    )
  })

  it("saves Word's empty result (five en spaces) written into a body field inside the field result", async () => {
    const { bytes, parsed, editor } = await open()
    const enSpaces = '\u2002'.repeat(5)
    const field = editor.state.doc.child(1)
    const { emptyLine: _placeholderLine, ...display } = field.attrs.fieldDisplay
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(editor.state.doc.child(0).nodeSize, undefined, {
        ...field.attrs,
        fieldDisplay: { ...display, left: enSpaces },
      }),
    )
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    editor.destroy()
    const savedXml = await documentXmlOf(await saveDocx(parsed, plan.saveBlocks))
    const sourceXml = await documentXmlOf(bytes)
    expect(fieldResult(savedXml)).toBe(enSpaces)
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
    const line = editor.state.doc.child(1)
    expect(line.attrs.fieldDisplay).toMatchObject({ kind: 'text', left: '', emptyLine: true })
    const { emptyLine: _placeholderLine, ...display } = line.attrs.fieldDisplay
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(editor.state.doc.child(0).nodeSize, undefined, {
        ...line.attrs,
        fieldDisplay: { ...display, left: 'Erika Mustermann' },
      }),
    )
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
