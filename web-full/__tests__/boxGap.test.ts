// @vitest-environment jsdom
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../../apps/docs/src/renderer/editor/convert'
import { editorExtensions } from '../../apps/docs/src/renderer/editor/extensions'
import { BOX_GAP_CLASS } from '../boxGap'
import { editScopePlugin } from '../editScope'
import { markersPlugin, setMarkers } from '../markersPlugin'
import { reviewModePlugin } from '../reviewMode'

const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
const HANGING = '<w:pPr><w:ind w:left="400" w:hanging="400"/></w:pPr>'
const legacyBox =
  '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Box"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="1"/></w:checkBox></w:ffData></w:fldChar></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
const formText =
  '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Text"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
const controlBox = `<w:sdt><w:sdtPr><w:id w:val="7"/><w14:checkbox ${W14}><w14:checked w14:val="0"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt>`
const plain = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const inCell = (inner: string) =>
  `<w:tbl><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>${inner}</w:tc></w:tr></w:tbl>`
const VALUE = '1233213'

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((editor) => editor.destroy()))

async function open(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  editor.registerPlugin(reviewModePlugin, (plugin, plugins) => [plugin, ...plugins])
  editors.push(editor)
  return { parsed, editor }
}

const gap = (editor: Editor) => editor.view.dom.querySelector<HTMLElement>(`.${BOX_GAP_CLASS}`)

function boxParagraph(editor: Editor) {
  let found = -1
  editor.state.doc.descendants((node, pos) => {
    if (found === -1 && node.isTextblock && /[☐☒]/.test(node.textContent)) found = pos
    return found === -1
  })
  return { pos: found, node: editor.state.doc.nodeAt(found)! }
}

const typeAtLineEnd = (editor: Editor, text: string) => {
  const { pos, node } = boxParagraph(editor)
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos + node.nodeSize - 1)),
  )
  editor.view.dispatch(editor.state.tr.insertText(text))
}

const savedXml = async (parsed: Awaited<ReturnType<typeof parseDocx>>, editor: Editor) => {
  const bytes = await saveDocx(
    parsed,
    pmDocToSavePlan(editor.getJSON() as never, parsed.blocks).saveBlocks,
  )
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
}

describe('a box on a hanging line with nothing after it', () => {
  it.each([
    [
      'a legacy box and a text field in a cell',
      inCell(`<w:p>${HANGING}${legacyBox}${formText}</w:p>`),
    ],
    ['a checkbox control alone', `<w:p>${HANGING}${controlBox}</w:p>`],
  ])('%s is drawn as wide as the hanging indent', async (_, body) => {
    const { editor } = await open(body)
    expect(gap(editor)?.style.minWidth).toBe('20pt')
    expect(gap(editor)?.textContent).toMatch(/^[☐☒]$/)
  })

  it.each([
    ['a tab', `<w:p>${HANGING}${legacyBox}<w:r><w:tab/></w:r>${formText}</w:p>`],
    ['a space and a label', `<w:p>${HANGING}${legacyBox}${plain(' Einzelbieter')}</w:p>`],
    ['no hanging indent', `<w:p>${legacyBox}${formText}</w:p>`],
  ])('with %s keeps the layout of the form', async (_, body) => {
    const { editor } = await open(body)
    expect(gap(editor)).toBeNull()
  })

  it('the gap stays while the text is typed', async () => {
    const { editor } = await open(inCell(`<w:p>${HANGING}${legacyBox}${formText}</w:p>`))
    typeAtLineEnd(editor, VALUE)
    expect(gap(editor)?.textContent).toBe('☒')
    expect(boxParagraph(editor).node.textContent).toBe(`☒${VALUE}`)
  })
})

describe('select-all in a field that opens with a box', () => {
  async function field(bodyXml: string) {
    const { parsed, editor } = await open(bodyXml)
    editor.registerPlugin(markersPlugin)
    editor.registerPlugin(editScopePlugin(vi.fn()), (plugin, plugins) => [plugin, ...plugins])
    setMarkers(editor.view, [
      {
        id: 'f',
        paragraph: editor.state.doc.child(0).attrs.docxPara,
        scope: 'paragraph',
        tone: 'proposed',
        label: 'F',
        actions: ['ask'],
      },
    ])
    editor.commands.setTextSelection(3)
    return { parsed, editor }
  }

  const pressSelectAll = (editor: Editor) =>
    editor.view.someProp('handleKeyDown', (handler) =>
      handler(editor.view, new KeyboardEvent('keydown', { key: 'a', metaKey: true })),
    )

  it.each([
    ['a legacy box', legacyBox],
    ['a checkbox control', controlBox],
  ])('selects the text after %s, so deleting it keeps the box', async (_, box) => {
    const { editor } = await field(`<w:p>${HANGING}${box}${plain(VALUE)}</w:p>`)
    expect(pressSelectAll(editor)).toBe(true)
    const { from, to } = editor.state.selection
    expect(editor.state.doc.textBetween(from, to)).toBe(VALUE)
    const glyph = editor.state.doc.child(0).textContent[0]
    editor.view.dispatch(editor.state.tr.deleteSelection())
    expect(editor.state.doc.child(0).textContent).toBe(glyph)
  })

  it('leaves the caret after the box when the field has no text', async () => {
    const { editor } = await field(`<w:p>${HANGING}${controlBox}</w:p>`)
    editor.commands.setTextSelection(2)
    expect(pressSelectAll(editor)).toBe(true)
    expect(editor.state.selection.empty).toBe(true)
    expect(editor.state.selection.from).toBe(2)
  })
})

describe('a box field with no text yet stays one line tall', () => {
  const trailingBreak = (editor: Editor) =>
    editor.view.dom.querySelector<HTMLElement>('.ProseMirror-trailingBreak')

  async function attached(bodyXml: string) {
    const { editor } = await open(bodyXml)
    document.body.append(editor.view.dom)
    const style = document.head.appendChild(document.createElement('style'))
    style.textContent = readFileSync(resolve(__dirname, '../host.css'), 'utf8')
    document.documentElement.classList.add('vound-review')
    return editor
  }

  afterEach(() => {
    document.head.replaceChildren()
    document.body.replaceChildren()
    document.documentElement.classList.remove('vound-review')
  })

  it('hides the break after the box, for a legacy box and a checkbox control', async () => {
    for (const box of [legacyBox, controlBox]) {
      const editor = await attached(`<w:p>${HANGING}${box}</w:p>`)
      expect(gap(editor)).not.toBeNull()
      expect(getComputedStyle(trailingBreak(editor)!).display).toBe('none')
    }
  })

  it('keeps the break of an empty paragraph with no box', async () => {
    const editor = await attached(`<w:p>${HANGING}</w:p>`)
    expect(getComputedStyle(trailingBreak(editor)!).display).not.toBe('none')
  })

  it('keeps the break that holds the new line after a line break', async () => {
    const editor = await attached(`<w:p>${HANGING}${controlBox}<w:r><w:br/></w:r></w:p>`)
    expect(getComputedStyle(trailingBreak(editor)!).display).not.toBe('none')
  })
})

describe('the gap never reaches the file', () => {
  it('a text field after a legacy box holds exactly the typed text, with no tab or space added', async () => {
    const { parsed, editor } = await open(inCell(`<w:p>${HANGING}${legacyBox}${formText}</w:p>`))
    typeAtLineEnd(editor, VALUE)
    const xml = await savedXml(parsed, editor)

    expect(xml).toContain(`<w:t xml:space="preserve">${VALUE}</w:t>`)
    expect(xml).not.toContain('<w:tab/>')
    expect(xml).not.toContain(' 1233213')
    expect(xml).not.toContain(BOX_GAP_CLASS)
  })

  it('text typed after a lone checkbox control has no tab or space added', async () => {
    const { parsed, editor } = await open(`<w:p>${HANGING}${controlBox}</w:p>`)
    typeAtLineEnd(editor, VALUE)
    const xml = await savedXml(parsed, editor)

    expect(xml).toContain(`<w:t xml:space="preserve">${VALUE}</w:t>`)
    expect(xml).not.toContain('<w:tab/>')
  })

  it('a save with no edit keeps the original paragraph', async () => {
    const line = `<w:p>${HANGING}${legacyBox}${formText}</w:p>`
    const { parsed, editor } = await open(line)
    expect(await savedXml(parsed, editor)).toContain(line)
  })
})
