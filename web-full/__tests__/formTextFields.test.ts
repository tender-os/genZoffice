// @vitest-environment jsdom
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../../apps/docs/src/renderer/editor/convert'
import { editorExtensions } from '../../apps/docs/src/renderer/editor/extensions'
import { editScopePlugin, setEditScope } from '../editScope'
import { formTextFieldAt, formTextPlugin } from '../formTextField'
import { markersPlugin, setMarkers } from '../markersPlugin'
import { reviewModePlugin } from '../reviewMode'

const EN_SPACE = ' '
const plain = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const TAB = '<w:r><w:tab/></w:r>'
const fieldOpen = (defaultText: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="F"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput>${defaultText ? `<w:default w:val="${defaultText}"/>` : ''}</w:textInput></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`
const FIELD_END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
const withDefault = (text: string) => `${fieldOpen(text)}${plain(text)}${FIELD_END}`
const BLANK = `${fieldOpen('')}${plain(EN_SPACE).repeat(5)}${FIELD_END}`
const withValue = (text: string) => `${fieldOpen('')}${plain(text)}${FIELD_END}`
const para = (inner: string) => `<w:p>${inner}</w:p>`

const LABELLED = para(`${plain('Firma:')}${TAB}${withDefault('Name eintragen')}`)
const THREE = para(
  `${withDefault('Tag')}${plain(', ')}${withDefault('Monat')}${TAB}${withDefault('Jahr')}`,
)
const BLANK_LINE = para(BLANK)
const BLANKS = para(`${BLANK}${plain(' / ')}${TAB}${plain('bis')}${BLANK}`)
const VALUE_LINE = para(`${plain('Ort: ')}${withValue('Muster GmbH')}`)
const FREE = para(plain('Fester Text des Formulars'))
const BODY = [LABELLED, THREE, BLANK_LINE, BLANKS, VALUE_LINE, FREE].join('')

const PROTECTION = '<w:documentProtection w:edit="forms" w:enforcement="1"/>'
const SETTINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${PROTECTION}</w:settings>`

const editors: Editor[] = []
afterEach(() => {
  editors.splice(0).forEach((editor) => editor.destroy())
  setEditScope('everything')
})

async function open(bodyXml: string, markedParagraphs: number[]) {
  const parsed = await parseDocx(
    await buildDocx({
      bodyXml,
      extraParts: [
        {
          path: 'word/settings.xml',
          xml: SETTINGS,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
        },
      ],
    }),
  )
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  const first = <T>(plugin: T, plugins: T[]) => [plugin, ...plugins]
  editor.registerPlugin(markersPlugin)
  editor.registerPlugin(reviewModePlugin, first)
  editor.registerPlugin(formTextPlugin, first)
  editor.registerPlugin(editScopePlugin(vi.fn()), first)
  setMarkers(
    editor.view,
    markedParagraphs.map((index) => ({
      id: `m${index}`,
      paragraph: editor.state.doc.child(index).attrs.docxPara,
      scope: 'paragraph' as const,
      tone: 'gap' as const,
      label: 'F',
      actions: ['ask' as const],
    })),
  )
  setEditScope('fields')
  editors.push(editor)
  return { parsed, editor }
}

/** the text input's own path: the plugins first, then the editor's default insertion */
function typeAt(editor: Editor, from: number, to: number, text: string) {
  const { view } = editor
  const handled = view.someProp('handleTextInput', (handler) =>
    handler(view, from, to, text, () => view.state.tr.insertText(text, from, to)),
  )
  if (handled) return
  view.dispatch(
    view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to)).insertText(text),
  )
}

const posOf = (editor: Editor, text: string) => {
  let found = -1
  editor.state.doc.descendants((node, pos) => {
    if (found === -1 && node.isText && node.text?.includes(text))
      found = pos + node.text.indexOf(text)
  })
  return found
}

const blankOf = (editor: Editor, block: number, occurrence: number) => {
  let start = editor.state.doc.resolve(0).posAtIndex(block) + 1
  const text = editor.state.doc.child(block).textContent
  let at = -1
  for (let count = 0; count <= occurrence; count += 1) {
    at = text.indexOf(EN_SPACE, count === 0 ? 0 : at + 5)
  }
  start += at
  return { from: start, to: start + 5 }
}

const saved = async (parsed: Awaited<ReturnType<typeof parseDocx>>, editor: Editor) => {
  const bytes = await saveDocx(
    parsed,
    pmDocToSavePlan(editor.getJSON() as never, parsed.blocks).saveBlocks,
  )
  const zip = await JSZip.loadAsync(bytes)
  return {
    xml: await zip.file('word/document.xml')!.async('string'),
    settings: await zip.file('word/settings.xml')!.async('string'),
  }
}

const paragraphsOf = (xml: string) => xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? []
const resultsOf = (paragraph: string) =>
  paragraph
    .split('w:fldCharType="separate"')
    .slice(1)
    .map((part) =>
      [...part.split('w:fldCharType="end"')[0].matchAll(/<w:t(?: [^>]*)?>([^<]*)/g)]
        .map((match) => match[1])
        .join(''),
    )
const fieldCount = (xml: string) => xml.split('FORMTEXT').length - 1

describe('a body line of legacy text form fields in a form protected for filling', () => {
  it('is text the caret can enter, not a protected block', async () => {
    const { editor, parsed } = await open(BODY, [0, 1])
    expect(parsed.protection?.edit).toBe('forms')
    const types: string[] = []
    editor.state.doc.forEach((node) => types.push(node.type.name))
    expect(types).not.toContain('docProtected')
  })

  it('writes typing over default text into the result, and keeps the label, the field and the protection', async () => {
    const { editor, parsed } = await open(BODY, [0, 1])
    const at = posOf(editor, 'eintragen')
    typeAt(editor, at, at, 'X')
    const { xml, settings } = await saved(parsed, editor)
    const line = paragraphsOf(xml)[0]
    expect(resultsOf(line)).toEqual(['X'])
    expect(line).toContain('>Firma:<')
    expect(fieldCount(xml)).toBe(fieldCount(BODY))
    expect(settings).toContain(PROTECTION)
  })

  it('keeps several fields of a line separate', async () => {
    const { editor, parsed } = await open(BODY, [0, 1])
    const at = posOf(editor, 'Monat')
    typeAt(editor, at, at, '1')
    const { xml } = await saved(parsed, editor)
    expect(resultsOf(paragraphsOf(xml)[1])).toEqual(['Tag', '1', 'Jahr'])
    expect(paragraphsOf(xml)[1]).toContain('>, <')
  })

  it('types inside a filled field and at its start, next to the label', async () => {
    const { editor, parsed } = await open(BODY, [4])
    const inside = posOf(editor, 'GmbH')
    typeAt(editor, inside, inside, 'X')
    const start = posOf(editor, 'Muster')
    typeAt(editor, start, start, 'Y')
    const { xml } = await saved(parsed, editor)
    expect(resultsOf(paragraphsOf(xml)[4])).toEqual(['YMuster XGmbH'])
    expect(paragraphsOf(xml)[4]).toContain('>Ort: <')
  })

  it('replaces a blank field whole, in a line of its own and between other fields', async () => {
    const { editor, parsed } = await open(BODY, [])
    const only = blankOf(editor, 2, 0)
    typeAt(editor, only.from + 2, only.from + 2, 'A')
    const second = blankOf(editor, 3, 1)
    typeAt(editor, second.from, second.from, 'B')
    const { xml } = await saved(parsed, editor)
    expect(resultsOf(paragraphsOf(xml)[2])).toEqual(['A'])
    expect(resultsOf(paragraphsOf(xml)[3])).toEqual([EN_SPACE.repeat(5), 'B'])
    expect(fieldCount(xml)).toBe(fieldCount(BODY))
  })

  it('locks the label of a marked line and text outside any field', async () => {
    const { editor, parsed } = await open(BODY, [0, 1])
    const label = posOf(editor, 'Firma') + 1
    typeAt(editor, label, label, 'X')
    const free = posOf(editor, 'Fester') + 1
    typeAt(editor, free, free, 'X')
    const separator = posOf(editor, ', ') + 1
    typeAt(editor, separator, separator, 'X')
    expect(editor.state.doc.textContent).not.toContain('X')
    const { xml } = await saved(parsed, editor)
    expect(paragraphsOf(xml)[0]).toContain('>Firma:<')
    expect(xml).toContain('>Fester Text des Formulars<')
  })

  it('keeps a field typed empty a field: the next text goes back into it, with no filler in the file', async () => {
    const { editor, parsed } = await open(BODY, [0, 1])
    const from = posOf(editor, 'Monat')
    editor.view.dispatch(
      editor.state.tr
        .setSelection(TextSelection.create(editor.state.doc, from, from + 5))
        .deleteSelection(),
    )
    typeAt(editor, from, editor.state.selection.to, '7')
    const { xml } = await saved(parsed, editor)
    expect(resultsOf(paragraphsOf(xml)[1])).toEqual(['Tag', '7', 'Jahr'])
    expect(xml).not.toContain('​')
  })

  it('keeps an emptied field in the file with an empty result', async () => {
    const { editor, parsed } = await open(BODY, [0, 1])
    const from = posOf(editor, 'Monat')
    editor.view.dispatch(
      editor.state.tr
        .setSelection(TextSelection.create(editor.state.doc, from, from + 5))
        .deleteSelection(),
    )
    const { xml } = await saved(parsed, editor)
    expect(resultsOf(paragraphsOf(xml)[1])).toEqual(['Tag', '', 'Jahr'])
    expect(fieldCount(xml)).toBe(fieldCount(BODY))
  })

  it('opens a saved empty field as a place to type again', async () => {
    const { editor, parsed } = await open(BODY, [0, 1])
    const from = posOf(editor, 'Monat')
    editor.view.dispatch(
      editor.state.tr
        .setSelection(TextSelection.create(editor.state.doc, from, from + 5))
        .deleteSelection(),
    )
    const bytes = await saveDocx(
      parsed,
      pmDocToSavePlan(editor.getJSON() as never, parsed.blocks).saveBlocks,
    )
    const reopened = await parseDocx(bytes)
    expect(reopened.blocks[1].runs?.filter((run) => run.formText === 1)).toMatchObject([
      { text: '\u200b', placeholder: true },
    ])
  })

  it('starts a hanging line that opens with a field at its indent', async () => {
    const { editor } = await open(
      para(`<w:pPr><w:ind w:left="400" w:hanging="400"/></w:pPr>${BLANK}`),
      [0],
    )
    const line = editor.view.dom.querySelector<HTMLElement>('p')
    expect(line?.style.textIndent).toMatch(/^0/)
  })

  it('leaves an unedited line as it was', async () => {
    const { editor, parsed } = await open(BODY, [0, 1])
    const { xml } = await saved(parsed, editor)
    expect(xml).toContain(LABELLED)
    expect(xml).toContain(THREE)
  })
})

describe('the shapes of legacy text fields a form puts in a body paragraph', () => {
  const BREAK = '<w:r><w:br/></w:r>'
  const ORPHAN_END = FIELD_END
  const AMOUNT = '5 % von der Vergabesumme'
  const SHAPES: Array<[string, string, (editor: Editor) => number, string[][]]> = [
    ['a blank field alone', para(BLANK), (editor) => blankOf(editor, 0, 0).from, [['A']]],
    [
      'a blank field before a line break',
      para(`${BLANK}${BREAK}`),
      (editor) => blankOf(editor, 0, 0).from,
      [['A']],
    ],
    [
      'two fields on two lines of one paragraph',
      para(`${BLANK}${BREAK}${withValue('Musterstraße 1')}`),
      (editor) => posOf(editor, 'straße'),
      [[EN_SPACE.repeat(5), 'MusterAstraße 1']],
    ],
    [
      'a field running on into the next paragraph, which then holds another',
      `${para(`${fieldOpen('')}${plain(AMOUNT)}${TAB}`)}${para(`${ORPHAN_END}${TAB.repeat(3)}${BLANK}`)}`,
      (editor) => posOf(editor, 'Vergabesumme'),
      [[`5 % von der AVergabesumme`], [EN_SPACE.repeat(5)]],
    ],
    [
      'a field after a label, running on into the next paragraph',
      `${para(`${plain('Betrag: ')}${fieldOpen('')}${plain(AMOUNT)}${TAB}`)}${para(`${ORPHAN_END}${TAB}${BLANK}`)}`,
      (editor) => posOf(editor, '5 %'),
      [[`A${AMOUNT}`], [EN_SPACE.repeat(5)]],
    ],
    [
      'two fields with a character between them',
      para(`${BLANK}${plain('/')}${BLANK}`),
      (editor) => blankOf(editor, 0, 1).from,
      [[EN_SPACE.repeat(5), 'A']],
    ],
  ]
  const count = (xml: string, element: string) => xml.split(element).length - 1

  it.each(SHAPES)(
    '%s: typing goes into its field, and the save keeps every field, break and tab',
    async (_, body, at, results) => {
      const { editor, parsed } = await open(body, [0])
      const types: string[] = []
      editor.state.doc.forEach((node) => types.push(node.type.name))
      expect(types).not.toContain('docProtected')
      const pos = at(editor)
      expect(formTextFieldAt(editor.state.doc, pos, pos)).not.toBeNull()
      typeAt(editor, pos, pos, 'A')
      const { xml } = await saved(parsed, editor)
      const paragraphs = paragraphsOf(xml).filter((paragraph) => paragraph.includes('fldChar'))
      expect(paragraphs.map(resultsOf)).toEqual(results)
      expect(xml).not.toMatch(/<w:t[^>]*>[^<]*\n/)
      for (const element of ['FORMTEXT', '<w:br/>', '<w:tab/>', 'w:fldCharType="end"']) {
        expect(count(xml, element)).toBe(count(body, element))
      }
    },
  )

  it('locks the tabs after a field that runs on into the next paragraph', async () => {
    const { editor } = await open(
      `${para(`${fieldOpen('')}${plain(AMOUNT)}${TAB}`)}${para(`${ORPHAN_END}${TAB.repeat(3)}${BLANK}`)}`,
      [0, 1],
    )
    const tabs = editor.state.doc.child(1).textContent.indexOf('\t')
    const at = editor.state.doc.resolve(0).posAtIndex(1) + 1 + tabs + 1
    typeAt(editor, at, at, 'X')
    expect(editor.state.doc.child(1).textContent).not.toContain('X')
  })
})
