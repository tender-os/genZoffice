// @vitest-environment jsdom
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../../apps/docs/src/renderer/editor/convert'
import { editorExtensions } from '../../apps/docs/src/renderer/editor/extensions'
import { placeholderAt } from '../placeholderText'

type Parsed = Awaited<ReturnType<typeof parseDocx>>

const plain = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const control = (id: string, text: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="${id}"/><w:text/></w:sdtPr><w:sdtContent>${plain(text)}</w:sdtContent></w:sdt>`
const paragraph = (inner: string) => `<w:p>${inner}</w:p>`
const inCell = (inner: string) =>
  `<w:tbl><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>${inner}</w:tc></w:tr></w:tbl>`

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((editor) => editor.destroy()))

function open(parsed: Parsed) {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
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

const save = (parsed: Parsed, editor: Editor) =>
  saveDocx(parsed, pmDocToSavePlan(editor.getJSON() as never, parsed.blocks).saveBlocks)

const reopen = async (bytes: Uint8Array) => {
  const parsed = await parseDocx(bytes)
  return { parsed, editor: open(parsed) }
}

const documentOf = async (bytes: Uint8Array) => {
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
  return new DOMParser().parseFromString(xml, 'application/xml')
}

const textIn = (element: Element) =>
  [...element.getElementsByTagName('w:t')].map((text) => text.textContent).join('')

/** the paragraph as parts: a literal's text, or a control's text in brackets */
async function parts(bytes: Uint8Array) {
  const dom = await documentOf(bytes)
  const line = dom.getElementsByTagName('w:p')[0]
  return [...line.children]
    .filter((child) => child.tagName === 'w:r' || child.tagName === 'w:sdt')
    .map((child) => (child.tagName === 'w:sdt' ? `[${textIn(child)}]` : textIn(child)))
    .join('')
}

function deleteText(editor: Editor, text: string) {
  const from = textPos(editor, text)
  editor.view.dispatch(editor.state.tr.delete(from, from + text.length))
  return from
}

const typeAt = (editor: Editor, at: number, text: string) => {
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at)))
  editor.view.dispatch(editor.state.tr.insertText(text))
}

/** the marked stand-in of an emptied control, at or after `from`; select it and type as a click would */
function fillEmptied(editor: Editor, from: number, text: string) {
  const range = placeholderAt(editor.state.doc, from)
  expect(range).not.toBeNull()
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, range!.from, range!.to)),
  )
  editor.view.dispatch(editor.state.tr.insertText(text))
}

const DATE = '09.11.2026'
const CASES = [
  {
    name: 'at the end of a paragraph after a label',
    body: () => paragraph(`${plain('Datum: ')}${control('1', DATE)}`),
    expected: (text: string) => `Datum: [${text}]`,
  },
  {
    name: 'in the middle of a paragraph',
    body: () => paragraph(`${plain('Am ')}${control('1', DATE)}${plain(' in Berlin')}`),
    expected: (text: string) => `Am [${text}] in Berlin`,
  },
  {
    name: 'at the start of a paragraph',
    body: () => paragraph(`${control('1', DATE)}${plain(' ist das Datum')}`),
    expected: (text: string) => `[${text}] ist das Datum`,
  },
]

const parsedOf = async (bodyXml: string) => parseDocx(await buildDocx({ bodyXml }))

describe('retyping an inline control', () => {
  it.each(CASES)(
    '$name: emptied and retyped before one save, the text stays inside',
    async ({ body, expected }) => {
      const parsed = await parsedOf(body())
      const editor = open(parsed)
      const at = deleteText(editor, DATE)
      typeAt(editor, at, '10.11.2026')

      expect(await parts(await save(parsed, editor))).toBe(expected('10.11.2026'))
    },
  )

  it.each(CASES)(
    '$name: emptied, saved, reopened, clicked and retyped, the text stays inside',
    async ({ body, expected }) => {
      const parsed = await parsedOf(body())
      const emptying = open(parsed)
      const at = deleteText(emptying, DATE)
      const emptied = await save(parsed, emptying)
      expect(await parts(emptied)).toBe(expected(''))

      const { parsed: again, editor } = await reopen(emptied)
      fillEmptied(editor, at, '10.11.2026')

      expect(await parts(await save(again, editor))).toBe(expected('10.11.2026'))
    },
  )

  it('two controls of one paragraph, each emptied and retyped', async () => {
    const body = paragraph(
      `${plain('Ort: ')}${control('1', 'Berlin')}${plain(', Datum: ')}${control('2', DATE)}`,
    )
    const parsed = await parsedOf(body)
    const emptying = open(parsed)
    deleteText(emptying, DATE)
    const ort = deleteText(emptying, 'Berlin')
    const emptied = await save(parsed, emptying)
    expect(await parts(emptied)).toBe('Ort: [], Datum: []')

    const { parsed: again, editor } = await reopen(emptied)
    const datum = textPos(editor, ', Datum: ') + ', Datum: '.length
    fillEmptied(editor, datum, '10.11.2026')
    fillEmptied(editor, ort, 'Hamburg')

    expect(await parts(await save(again, editor))).toBe('Ort: [Hamburg], Datum: [10.11.2026]')
  })

  it('a value replaced whole, and typing at either end of it, stay inside', async () => {
    const body = paragraph(`${plain('Datum: ')}${control('1', DATE)}${plain(' Uhr')}`)
    const parsed = await parsedOf(body)
    const editor = open(parsed)
    const at = textPos(editor, DATE)
    typeAt(editor, at + DATE.length, '!')
    typeAt(editor, at, '>')

    expect(await parts(await save(parsed, editor))).toBe(`Datum: [>${DATE}!] Uhr`)
  })

  it('an emptied control in a table cell can be clicked and filled again', async () => {
    const parsed = await parsedOf(inCell(paragraph(control('1', DATE))))
    const emptying = open(parsed)
    const at = deleteText(emptying, DATE)
    const { parsed: again, editor } = await reopen(await save(parsed, emptying))
    fillEmptied(editor, at, '10.11.2026')

    const dom = await documentOf(await save(again, editor))
    expect([...dom.getElementsByTagName('w:sdt')].map(textIn)).toEqual(['10.11.2026'])
  })

  it('a file whose text already sits before an empty control can fill that control again', async () => {
    const broken = paragraph(`${plain('Datum: 10.11.2026')}${control('1', '')}`)
    const parsed = await parsedOf(broken)
    const editor = open(parsed)
    const end = textPos(editor, '10.11.2026') + '10.11.2026'.length
    fillEmptied(editor, end, '11.11.2026')

    expect(await parts(await save(parsed, editor))).toBe('Datum: 10.11.2026[11.11.2026]')
  })

  it.each([
    ['in the body', (inner: string) => inner],
    ['in a table cell', inCell],
  ])('editing beside an emptied control %s writes no stand-in into the file', async (_, place) => {
    const body =
      place(paragraph(`${plain('Name: ')}${control('1', '')}`)) + paragraph(plain('Sonst'))
    const parsed = await parsedOf(body)
    const editor = open(parsed)
    const at = textPos(editor, 'Name')
    typeAt(editor, at, 'Vor')
    const xml = await (
      await JSZip.loadAsync(await save(parsed, editor))
    )
      .file('word/document.xml')!
      .async('string')

    expect(xml).not.toContain('​')
    expect(xml).toContain('VorName')
  })
})
