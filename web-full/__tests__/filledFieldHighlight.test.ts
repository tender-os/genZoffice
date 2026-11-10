// @vitest-environment jsdom
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../../apps/docs/src/renderer/editor/convert'
import { editorExtensions } from '../../apps/docs/src/renderer/editor/extensions'
import { settleField } from '../fieldSettle'
import { applyFieldValue } from '../fieldWrite'
import { filledFieldHighlightPlugin, formTextPlugin } from '../formTextField'

const MARKED = '<w:rPr><w:highlight w:val="yellow"/></w:rPr>'
const marked = (inner: string) => `<w:r>${MARKED}${inner}</w:r>`
const DEFAULT_TEXT = 'Bitte Name angeben'
const FIELD =
  marked(
    `<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val=""/><w:enabled/><w:textInput><w:default w:val="${DEFAULT_TEXT}"/></w:textInput></w:ffData></w:fldChar>`,
  ) +
  marked('<w:instrText xml:space="preserve"> FORMTEXT </w:instrText>') +
  marked('<w:fldChar w:fldCharType="separate"/>') +
  marked(`<w:t>${DEFAULT_TEXT}</w:t>`) +
  marked('<w:fldChar w:fldCharType="end"/>')
const LINE = `<w:p><w:r><w:t xml:space="preserve">Mitglied </w:t></w:r>${FIELD}</w:p>`
const fieldWith = (result: string) => FIELD.replace(`<w:t>${DEFAULT_TEXT}</w:t>`, result)
const BLANK_LINE = `<w:p><w:r><w:t xml:space="preserve">Mitglied </w:t></w:r>${fieldWith('<w:t xml:space="preserve"></w:t>')}</w:p>`
const AI_VALUE_LINE = `<w:p><w:r><w:t xml:space="preserve">Mitglied </w:t></w:r>${fieldWith('<w:t>Erika Muster</w:t>')}</w:p>`
const NOTE = `<w:p>${marked('<w:t>Hinweis</w:t>')}</w:p>`

const editors: Editor[] = []
afterEach(() => {
  editors.splice(0).forEach((editor) => editor.destroy())
  document.head.innerHTML = ''
  document.body.innerHTML = ''
})

async function open() {
  const parsed = await parseDocx(
    await buildDocx({ bodyXml: LINE + LINE + NOTE + BLANK_LINE + AI_VALUE_LINE }),
  )
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor
    .chain()
    .setMeta('addToHistory', false)
    .setContent(blocksToPmDoc(parsed.blocks) as never)
    .run()
  editor.registerPlugin(formTextPlugin)
  editor.registerPlugin(filledFieldHighlightPlugin)
  editors.push(editor)
  return { parsed, editor }
}

/** the highlight of every text node of a paragraph, by its text */
const highlightsOf = (editor: Editor, paragraph: number) => {
  const found: Record<string, unknown> = {}
  editor.state.doc.child(paragraph).forEach((node) => {
    const style = node.marks.find((mark) => mark.type.name === 'docTextStyle')
    found[node.text ?? ''] = style?.attrs.highlight ?? null
  })
  return found
}

const writeHost = (editor: Editor, paragraph: number, value: string) =>
  applyFieldValue(editor.view, {
    block: { paragraph: editor.state.doc.child(paragraph).attrs.docxPara, scope: 'paragraph' },
    previous: DEFAULT_TEXT,
    value,
  })

const savedParagraphs = async (parsed: Awaited<ReturnType<typeof parseDocx>>, editor: Editor) => {
  const bytes = await saveDocx(
    parsed,
    pmDocToSavePlan(editor.getJSON() as never, parsed.blocks).saveBlocks,
  )
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
  return xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? []
}
const highlightCount = (xml: string) => xml.split('<w:highlight').length - 1

describe("a legacy text field the form's author marked with highlight", () => {
  it('loses it once the host writes a value, and saves without it; open fields and other text keep theirs', async () => {
    const { parsed, editor } = await open()
    expect(writeHost(editor, 0, 'Erika Muster')).toEqual(expect.objectContaining({ ok: true }))
    expect(highlightsOf(editor, 0)).toEqual({ 'Mitglied ': null, 'Erika Muster': null })
    expect(highlightsOf(editor, 1)).toEqual({ 'Mitglied ': null, [DEFAULT_TEXT]: 'yellow' })
    const [filled, untouched, note] = await savedParagraphs(parsed, editor)
    expect(filled).toContain('>Erika Muster<')
    expect(highlightCount(filled)).toBe(0)
    expect(highlightCount(untouched)).toBe(5)
    expect(highlightCount(note)).toBe(1)
  })

  it('loses it when the reviewer types the value', async () => {
    const { parsed, editor } = await open()
    let from = 0
    editor.state.doc.descendants((node, pos) => {
      if (from === 0 && node.text === DEFAULT_TEXT) from = pos
    })
    const { view } = editor
    const to = from + DEFAULT_TEXT.length
    const typed = view.someProp('handleTextInput', (handler) =>
      handler(view, from, to, 'Erika Muster', () =>
        view.state.tr.insertText('Erika Muster', from, to),
      ),
    )
    expect(typed).toBe(true)
    expect(highlightsOf(editor, 0)).toEqual({ 'Mitglied ': null, 'Erika Muster': null })
    expect(highlightCount((await savedParagraphs(parsed, editor))[0])).toBe(0)
  })

  it('gets it back with the undo of the value', async () => {
    const { parsed, editor } = await open()
    writeHost(editor, 0, 'Erika Muster')
    editor.commands.undo()
    expect(highlightsOf(editor, 0)).toEqual({ 'Mitglied ': null, [DEFAULT_TEXT]: 'yellow' })
    expect(highlightCount((await savedParagraphs(parsed, editor))[0])).toBe(5)
  })
})

describe('a highlighted field the host settles (accepted, or marked blank)', () => {
  const settle = (editor: Editor, paragraph: number) =>
    settleField(editor.view, {
      paragraph: editor.state.doc.child(paragraph).attrs.docxPara,
      scope: 'paragraph',
    })

  it.each([
    ['marked blank', 3],
    ['an AI value accepted without an edit', 4],
  ])('saves without it when %s, and gets it back with an undo', async (_case, paragraph) => {
    const { parsed, editor } = await open()
    settle(editor, paragraph)
    expect(Object.values(highlightsOf(editor, paragraph))).not.toContain('yellow')
    expect(highlightCount((await savedParagraphs(parsed, editor))[paragraph])).toBe(0)
    editor.commands.undo()
    expect(highlightCount((await savedParagraphs(parsed, editor))[paragraph])).toBe(5)
  })

  it('leaves the other fields and the text outside fields as they were', async () => {
    const { parsed, editor } = await open()
    settle(editor, 3)
    const saved = await savedParagraphs(parsed, editor)
    expect(saved.map(highlightCount)).toEqual([5, 5, 1, 0, 5])
  })
})
