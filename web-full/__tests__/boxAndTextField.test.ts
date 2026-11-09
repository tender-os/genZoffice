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
import { placeCaret } from '../fieldWrite'
import { markersPlugin, setMarkers } from '../markersPlugin'
import { placeholderAt } from '../placeholderText'
import { reviewModePlugin } from '../reviewMode'

type Parsed = Awaited<ReturnType<typeof parseDocx>>
const EMPTY_RESULT = '     '

const box = (checked: 0 | 1) =>
  `<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Box"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="${checked}"/></w:checkBox></w:ffData></w:fldChar></w:r>` +
  '<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
const formText = (result: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Text"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`
const blankResult = `<w:r><w:t xml:space="preserve">${EMPTY_RESULT}</w:t></w:r>`
const TAB = '<w:r><w:tab/></w:r>'
const label = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const paragraph = (inner: string) => `<w:p>${inner}</w:p>`
const inCell = (inner: string) =>
  `<w:tbl><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>${inner}</w:tc></w:tr></w:tbl>`
const VALUE = '1233213'

const editors: Editor[] = []
afterEach(() => {
  editors.splice(0).forEach((editor) => editor.destroy())
  setEditScope('everything')
})

async function open(bodyXml: string, scope: 'everything' | 'fields') {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  const first = <T>(plugin: T, plugins: T[]) => [plugin, ...plugins]
  editor.registerPlugin(markersPlugin)
  editor.registerPlugin(reviewModePlugin, first)
  editor.registerPlugin(editScopePlugin(vi.fn()), first)
  editors.push(editor)
  setEditScope(scope)
  return { parsed, editor }
}

const fieldParagraph = (editor: Editor) => {
  let found = -1
  editor.state.doc.descendants((node, pos) => {
    if (found === -1 && node.isTextblock && node.textContent !== '') found = pos
    return found === -1
  })
  return { pos: found, node: editor.state.doc.nodeAt(found)! }
}

/** types as the browser does: the view's text input handlers, else a plain insertion at the selection */
function type(editor: Editor, text: string) {
  const { view } = editor
  const { from, to } = view.state.selection
  const handled = view.someProp('handleTextInput', (handler) =>
    handler(view, from, to, text, () => view.state.tr.insertText(text, from, to)),
  )
  if (!handled) view.dispatch(view.state.tr.insertText(text, from, to))
}

/** a click in the line's empty area: the caret at the end of the line */
function clickLineEnd(editor: Editor) {
  const { pos, node } = fieldParagraph(editor)
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos + node.nodeSize - 1)),
  )
}

const save = async (parsed: Parsed, editor: Editor) => {
  const bytes = await saveDocx(
    parsed,
    pmDocToSavePlan(editor.getJSON() as never, parsed.blocks).saveBlocks,
  )
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
  return new DOMParser().parseFromString(xml, 'application/xml')
}

/** the w:t texts between a field's separate and end, field by field */
function fieldResults(dom: Document, code: string) {
  const runs = [...dom.getElementsByTagName('w:r')]
  const results: string[] = []
  runs.forEach((run, index) => {
    if (!run.getElementsByTagName('w:instrText')[0]?.textContent?.includes(code)) return
    let text = ''
    for (let at = index + 1; at < runs.length; at += 1) {
      const type = runs[at].getElementsByTagName('w:fldChar')[0]?.getAttribute('w:fldCharType')
      if (type === 'end') break
      text += [...runs[at].getElementsByTagName('w:t')].map((t) => t.textContent).join('')
    }
    results.push(text)
  })
  return results
}

const boxDefaults = (dom: Document) =>
  [...dom.getElementsByTagName('w:default')].map((node) => node.getAttribute('w:val'))

describe.each(['everything', 'fields'] as const)('a box and a text field, editing %s', (scope) => {
  const markFields = (editor: Editor) => {
    const { node } = fieldParagraph(editor)
    setMarkers(editor.view, [
      {
        id: 's31',
        paragraph: node.attrs.docxPara,
        scope: 'cell',
        tone: 'gap',
        label: 'Register',
        actions: [],
      },
    ])
  }

  it.each([
    ['a text field with no result', paragraph(`${box(1)}${formText('')}`)],
    ['a tab and a text field with no result', paragraph(`${box(1)}${TAB}${formText('')}`)],
    ["a text field showing Word's blank", paragraph(`${box(1)}${formText(blankResult)}`)],
    ["a tab and Word's blank", paragraph(`${box(1)}${TAB}${formText(blankResult)}`)],
  ])('in a cell, after a ticked box: %s takes the typing', async (_, line) => {
    const { parsed, editor } = await open(inCell(line), scope)
    markFields(editor)
    clickLineEnd(editor)
    for (const key of VALUE) type(editor, key)

    expect(fieldParagraph(editor).node.textContent.startsWith('☒')).toBe(true)
    expect(fieldParagraph(editor).node.textContent.endsWith(VALUE)).toBe(true)
    expect(fieldParagraph(editor).node.textContent).not.toContain(' ')
    const dom = await save(parsed, editor)
    expect(fieldResults(dom, 'FORMTEXT')).toEqual([VALUE])
    expect(fieldResults(dom, 'FORMCHECKBOX')).toEqual([''])
    expect(dom.getElementsByTagName('w:ffData')).toHaveLength(2)
    expect(boxDefaults(dom)).toEqual(['1'])
  })

  it("the caret placed on the field selects Word's blank, so typing replaces it", async () => {
    const { parsed, editor } = await open(
      inCell(paragraph(`${box(0)}${TAB}${formText(blankResult)}`)),
      scope,
    )
    markFields(editor)
    const { node } = fieldParagraph(editor)
    Element.prototype.scrollIntoView = vi.fn()
    expect(placeCaret(editor.view, { paragraph: node.attrs.docxPara, scope: 'cell' })).toBe(true)
    const { from, to } = editor.state.selection
    expect(editor.state.doc.textBetween(from, to)).toBe(EMPTY_RESULT)
    type(editor, VALUE)

    const dom = await save(parsed, editor)
    expect(fieldResults(dom, 'FORMTEXT')).toEqual([VALUE])
    expect(boxDefaults(dom)).toEqual(['0'])
  })

  it('an inline control after a box takes the typing inside it and keeps the box', async () => {
    const control =
      '<w:sdt><w:sdtPr><w:id w:val="4"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Text eingeben</w:t></w:r></w:sdtContent></w:sdt>'
    const { parsed, editor } = await open(inCell(paragraph(`${box(1)}${TAB}${control}`)), scope)
    markFields(editor)
    const { pos, node } = fieldParagraph(editor)
    const range = placeholderAt(editor.state.doc, pos + node.nodeSize - 1)!
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, range.from, range.to)),
    )
    type(editor, VALUE)

    const dom = await save(parsed, editor)
    expect(dom.getElementsByTagName('w:sdtContent')[0].textContent).toBe(VALUE)
    expect(boxDefaults(dom)).toEqual(['1'])
    expect(dom.getElementsByTagName('w:ffData')).toHaveLength(1)
  })
})

describe('a box, a label and a text field in a body paragraph', () => {
  it.each([
    ["Word's blank", blankResult],
    ['no result', ''],
  ])('text added after the label lands in the field showing %s', async (_, result) => {
    const line = paragraph(`${box(1)}${label(' Handelsregister: ')}${formText(result)}`)
    const { parsed, editor } = await open(line, 'everything')
    const field = editor.state.doc.child(0)
    const left = `${field.attrs.fieldDisplay.left} HRB 42`
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, {
        ...field.attrs,
        fieldDisplay: { ...field.attrs.fieldDisplay, left },
      }),
    )

    const dom = await save(parsed, editor)
    expect(fieldResults(dom, 'FORMTEXT')).toEqual(['HRB 42'])
    expect(boxDefaults(dom)).toEqual(['1'])
    const literal = [...dom.getElementsByTagName('w:p')[0].children].filter((child) =>
      child.textContent?.includes('Handelsregister'),
    )
    expect(literal.map((child) => child.textContent)).toEqual([' Handelsregister: '])
  })
})
