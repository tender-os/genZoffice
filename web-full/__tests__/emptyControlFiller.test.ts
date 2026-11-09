// @vitest-environment jsdom
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import { EMPTY_CONTROL_TEXT, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../../apps/docs/src/renderer/editor/convert'
import { editorExtensions } from '../../apps/docs/src/renderer/editor/extensions'
import { rewriteKeepingHistory } from '../../apps/docs/src/renderer/editor/rewrite-keeping-history'
import { blockText } from '../fieldWrite'
import { placeholderAt } from '../placeholderText'

type Parsed = Awaited<ReturnType<typeof parseDocx>>

const plain = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const control = (text: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="1"/><w:text/></w:sdtPr><w:sdtContent>${plain(text)}</w:sdtContent></w:sdt>`
const LINE = `<w:p>${plain('Datum: ')}${control('09.11.2026')}</w:p>`

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

const xmlOf = async (bytes: Uint8Array) =>
  (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')

const wtsOf = async (bytes: Uint8Array, within: string) => {
  const dom = new DOMParser().parseFromString(await xmlOf(bytes), 'application/xml')
  const root = dom.getElementsByTagName(within)[0]
  return [...root.getElementsByTagName('w:t')].map((text) => text.textContent)
}

/** an opened document whose control an earlier save emptied: the editor shows the filler */
async function emptied() {
  const parsed = await parseDocx(await buildDocx({ bodyXml: LINE }))
  const emptying = open(parsed)
  const from = textPos(emptying, '09.11.2026')
  emptying.view.dispatch(emptying.state.tr.delete(from, from + 10))
  const again = await parseDocx(await save(parsed, emptying))
  const editor = open(again)
  const range = placeholderAt(editor.state.doc, textPos(editor, EMPTY_CONTROL_TEXT))!
  return { parsed: again, editor, range }
}

const VALUE = '10.11.2026'

describe('the filler of an emptied control never reaches the file', () => {
  it('type, save, empty, save, type, save in one open editor keeps the exact text', async () => {
    let parsed = await parseDocx(await buildDocx({ bodyXml: LINE }))
    const editor = open(parsed)
    /** a save, and the editor taking over the saved file's parse, as an embedded save does without a reload */
    const saveAndAdopt = async () => {
      const bytes = await save(parsed, editor)
      parsed = await parseDocx(bytes)
      const tr = editor.state.tr
      const next = editor.schema.nodeFromJSON(blocksToPmDoc(parsed.blocks))
      expect(rewriteKeepingHistory(tr, next)).toBe(true)
      if (tr.docChanged) editor.view.dispatch(tr.setMeta('addToHistory', false))
      return bytes
    }
    const from = textPos(editor, '09.11.2026')
    editor.view.dispatch(editor.state.tr.insertText('12.12.2026', from, from + 10))
    expect(await wtsOf(await saveAndAdopt(), 'w:sdt')).toEqual(['12.12.2026'])

    editor.view.dispatch(editor.state.tr.delete(from, from + 10))
    editor.commands.setTextSelection(from)
    expect(await wtsOf(await saveAndAdopt(), 'w:sdt')).toEqual([''])
    expect(editor.state.doc.textContent).toBe(`Datum: ${EMPTY_CONTROL_TEXT}`)

    for (const key of VALUE) editor.view.dispatch(editor.state.tr.insertText(key))
    expect(await wtsOf(await save(parsed, editor), 'w:sdt')).toEqual([VALUE])
    expect(editor.state.doc.textContent).toBe(`Datum: ${VALUE}`)
  })

  it('typing at its left edge', async () => {
    const { parsed, editor, range } = await emptied()
    editor.view.dispatch(editor.state.tr.insertText(VALUE, range.from))
    expect(await wtsOf(await save(parsed, editor), 'w:sdt')).toEqual([VALUE])
  })

  it('typing at its right edge, one key at a time', async () => {
    const { parsed, editor, range } = await emptied()
    editor.commands.setTextSelection(range.to)
    for (const key of VALUE) editor.view.dispatch(editor.state.tr.insertText(key))
    expect(await wtsOf(await save(parsed, editor), 'w:sdt')).toEqual([VALUE])
    expect(editor.state.doc.textContent).not.toContain(EMPTY_CONTROL_TEXT)
  })

  it('text the browser wrote into its span, read back with the mark', async () => {
    const { parsed, editor, range } = await emptied()
    const marked = editor.state.schema.text(
      `${EMPTY_CONTROL_TEXT}${VALUE}`,
      editor.state.doc.nodeAt(range.from)!.marks,
    )
    editor.view.dispatch(editor.state.tr.replaceWith(range.from, range.to, marked))
    expect(await wtsOf(await save(parsed, editor), 'w:sdt')).toEqual([VALUE])
  })

  it('a filler that lost its mark is still dropped when the control takes text', async () => {
    const { parsed, editor, range } = await emptied()
    const bare = editor.state.schema.text(`${EMPTY_CONTROL_TEXT}${VALUE}`)
    const change = editor.state.tr.replaceWith(range.from, range.to, bare).setMeta('history$', {})
    editor.view.dispatch(change)
    expect(await wtsOf(await save(parsed, editor), 'w:sdt')).toEqual([VALUE])
  })

  it('reports the field text to the host without the filler', async () => {
    const { editor, range } = await emptied()
    const paragraph = editor.state.doc.resolve(range.from).parent
    expect(blockText(paragraph, range.from - 1)).toBe('Datum: ')
  })
})

describe("a form's own zero-width characters stay", () => {
  it('in a control and a label of the edited paragraph', async () => {
    const own = `<w:p>${plain('Da​tum: ')}${control('A​B')}${plain(' Uhr')}</w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: own }))
    const editor = open(parsed)
    editor.view.dispatch(editor.state.tr.insertText('!', textPos(editor, ' Uhr') + 4))
    const saved = await save(parsed, editor)
    expect(await wtsOf(saved, 'w:p')).toEqual(['Da​tum: ', 'A​B', ' Uhr!'])
  })

  it('in a label right before an emptied control that takes text', async () => {
    const own = `<w:p>${plain('Datum:​')}${control('')}</w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: own }))
    const editor = open(parsed)
    const range = placeholderAt(editor.state.doc, textPos(editor, EMPTY_CONTROL_TEXT) + 1)!
    editor.view.dispatch(editor.state.tr.insertText(VALUE, range.to))
    const saved = await save(parsed, editor)
    expect(await wtsOf(saved, 'w:p')).toEqual(['Datum:​', VALUE])
  })
})
