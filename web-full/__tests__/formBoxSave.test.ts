// @vitest-environment jsdom
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { EditorView } from '@tiptap/pm/view'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../../apps/docs/src/renderer/editor/convert'
import { editorExtensions } from '../../apps/docs/src/renderer/editor/extensions'
import { rewriteKeepingHistory } from '../../apps/docs/src/renderer/editor/rewrite-keeping-history'
import { boxTickPlugin, toggleBoxAt } from '../boxTick'
import { editScopePlugin } from '../editScope'

type Parsed = Awaited<ReturnType<typeof parseDocx>>

const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
const legacyBox = (checked: 0 | 1) =>
  `<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Box"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="${checked}"/></w:checkBox></w:ffData></w:fldChar></w:r>` +
  '<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
const formText = (result: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Text"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`
const BLANK = '<w:r><w:t xml:space="preserve">     </w:t></w:r>'
const plain = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const controlBox = (checked: 0 | 1, glyph: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="7"/><w14:checkbox ${W14}><w14:checked w14:val="${checked}"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr>` +
  `<w:sdtContent><w:r><w:t>${glyph}</w:t></w:r></w:sdtContent></w:sdt>`
const paragraph = (inner: string) => `<w:p>${inner}</w:p>`
const inCell = (inner: string) =>
  `<w:tbl><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>${inner}</w:tc></w:tr></w:tbl>`

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((editor) => editor.destroy()))

async function open(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  editor.registerPlugin(editScopePlugin(vi.fn()), (plugin, plugins) => [plugin, ...plugins])
  editors.push(editor)
  return { parsed, editor }
}

const glyphPos = (editor: Editor) => {
  let found = -1
  editor.state.doc.descendants((node, pos) => {
    if (found === -1 && node.isText && /[☐☒]/.test(node.text ?? ''))
      found = pos + (node.text ?? '').search(/[☐☒]/)
  })
  return found
}

const toggle = (editor: Editor) =>
  editor.view.dispatch(toggleBoxAt(editor.state, glyphPos(editor))!)

const typeAfterGlyph = (editor: Editor, text: string) => {
  editor.commands.setTextSelection(glyphPos(editor) + 1)
  editor.view.dispatch(editor.state.tr.insertText(text))
}

const typeAtLineEnd = (editor: Editor, text: string) => {
  editor.commands.setTextSelection(editor.state.doc.resolve(glyphPos(editor)).end())
  editor.view.dispatch(editor.state.tr.insertText(text))
}

const saved = async (parsed: Parsed, editor: Editor) => {
  const bytes = await saveDocx(
    parsed,
    pmDocToSavePlan(editor.getJSON() as never, parsed.blocks).saveBlocks,
  )
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
  return { bytes, xml, dom: new DOMParser().parseFromString(xml, 'application/xml') }
}

const fields = (xml: string, code: string) => xml.split(code).length - 1
const checkedOf = (dom: Document) =>
  dom.getElementsByTagName('w:checked')[0]?.getAttribute('w:val') ?? null
const controlTexts = (dom: Document) =>
  [...dom.getElementsByTagName('w:sdtContent')].map((content) => content.textContent)
const controlChecked = (dom: Document) =>
  dom.getElementsByTagName('w14:checked')[0].getAttribute('w14:val')

describe('a form field survives every edit of its paragraph', () => {
  it.each([
    ['no result', ''],
    ["Word's blank", BLANK],
  ])(
    'a tick beside a text field with %s keeps the field and saves the state',
    async (_, result) => {
      const { parsed, editor } = await open(inCell(paragraph(`${legacyBox(1)}${formText(result)}`)))
      toggle(editor)
      const { xml, dom } = await saved(parsed, editor)

      expect(fields(xml, 'FORMTEXT')).toBe(1)
      expect(fields(xml, 'FORMCHECKBOX')).toBe(1)
      expect(checkedOf(dom)).toBe('0')
    },
  )

  it('a tick and typed text in one save both land', async () => {
    const { parsed, editor } = await open(inCell(paragraph(`${legacyBox(0)}${formText('')}`)))
    toggle(editor)
    typeAtLineEnd(editor, 'HRB 42')
    const { xml, dom } = await saved(parsed, editor)

    expect(fields(xml, 'FORMTEXT')).toBe(1)
    expect(checkedOf(dom)).toBe('1')
    expect(xml).toContain('HRB 42')
  })

  it('a paragraph with a form field takes no Enter', async () => {
    const { parsed, editor } = await open(
      inCell(paragraph(`${legacyBox(1)}${plain(' Text ')}${formText(BLANK)}`)),
    )
    typeAtLineEnd(editor, '')
    editor.commands.enter()
    const { xml } = await saved(parsed, editor)

    expect(fields(xml, 'FORMTEXT')).toBe(1)
    expect(fields(xml, 'FORMCHECKBOX')).toBe(1)
  })
})

describe('a table of form fields across saves in one open editor', () => {
  const row = (inner: string) =>
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>${paragraph(inner)}</w:tc></w:tr>`
  const control =
    '<w:sdt><w:sdtPr><w:id w:val="4"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Wert</w:t></w:r></w:sdtContent></w:sdt>'
  const TABLE =
    '<w:tbl><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>' +
    row(`${legacyBox(1)}${formText('')}`) +
    row(`${legacyBox(1)}${formText(BLANK)}`) +
    row(`${legacyBox(1)}${control}`) +
    row(plain('Sonst')) +
    '</w:tbl>'

  it('a table whose run formatting differs from the saved parse keeps every field', async () => {
    const opened = await open(TABLE)
    let parsed = opened.parsed
    const { editor } = opened
    typeAtLineEnd(editor, 'HRB1')
    toggle(editor)
    parsed = await parseDocx((await saved(parsed, editor)).bytes)
    const tr = editor.state.tr
    expect(
      rewriteKeepingHistory(tr, editor.schema.nodeFromJSON(blocksToPmDoc(parsed.blocks))),
    ).toBe(true)
    if (tr.docChanged) editor.view.dispatch(tr.setMeta('addToHistory', false))
    let end = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.isTextblock && node.textContent === 'Sonst') end = pos + node.nodeSize - 1
      return end === -1
    })
    editor.commands.setTextSelection(end)
    editor.view.dispatch(editor.state.tr.insertText(' mehr'))
    const { xml, dom } = await saved(parsed, editor)

    expect(fields(xml, 'FORMTEXT')).toBe(2)
    expect(fields(xml, 'FORMCHECKBOX')).toBe(3)
    expect(fields(xml, '<w:sdt>')).toBe(1)
    expect(dom.getElementsByTagName('w:checked')[0].getAttribute('w:val')).toBe('0')
    expect(xml).toContain('Sonst mehr')
  })
})

describe('a body paragraph with a box, a label and a text field', () => {
  const line = paragraph(`${legacyBox(1)}${plain(' Handelsregister: ')}${formText(BLANK)}`)

  it('shows the box and the label as text', async () => {
    const { editor } = await open(line)
    expect(editor.state.doc.child(0).type.name).toBe('docParagraph')
    expect(editor.state.doc.child(0).textContent.startsWith('☒ Handelsregister: ')).toBe(true)
  })

  it('takes typing in the field and keeps the box', async () => {
    const { parsed, editor } = await open(line)
    typeAtLineEnd(editor, 'HRB 42')
    const { xml, dom } = await saved(parsed, editor)

    expect(xml).toContain('>     HRB 42<')
    expect(fields(xml, 'FORMTEXT')).toBe(1)
    expect(checkedOf(dom)).toBe(null)
    expect(dom.getElementsByTagName('w:default')[0].getAttribute('w:val')).toBe('1')
  })

  it('a tick is saved in the field', async () => {
    const { parsed, editor } = await open(line)
    toggle(editor)
    const { xml, dom } = await saved(parsed, editor)

    expect(fields(xml, 'FORMTEXT')).toBe(1)
    expect(checkedOf(dom)).toBe('0')
  })
})

describe('a checkbox content control', () => {
  const labelled = paragraph(`${controlBox(0, '☐')}${plain('\tEinzelbieter')}`)
  const alone = paragraph(controlBox(0, '☐'))

  it('a tick saves the glyph and w14:checked together', async () => {
    const { parsed, editor } = await open(labelled)
    toggle(editor)
    const { dom } = await saved(parsed, editor)

    expect(controlTexts(dom)).toEqual(['☒'])
    expect(controlChecked(dom)).toBe('1')
  })

  it('text typed right after it stays outside it, in the label', async () => {
    const { parsed, editor } = await open(labelled)
    typeAfterGlyph(editor, '12')
    const { dom } = await saved(parsed, editor)

    expect(controlTexts(dom)).toEqual(['☐'])
    expect(controlChecked(dom)).toBe('0')
    expect(dom.getElementsByTagName('w:p')[0].textContent?.replace(/\s/g, '')).toBe(
      '☐12Einzelbieter',
    )
  })

  it('deleting the label and typing keep the control as it was', async () => {
    const { parsed, editor } = await open(labelled)
    const from = glyphPos(editor) + 1
    editor.view.dispatch(editor.state.tr.delete(from + 1, from + 1 + 'Einzelbieter'.length))
    editor.commands.setTextSelection(from + 1)
    editor.view.dispatch(editor.state.tr.insertText('99'))
    const { dom } = await saved(parsed, editor)

    expect(controlTexts(dom)).toEqual(['☐'])
    expect(controlChecked(dom)).toBe('0')
  })

  it('alone in its paragraph, text typed after it goes after the control', async () => {
    const { parsed, editor } = await open(alone)
    typeAfterGlyph(editor, '12')
    const { xml, dom } = await saved(parsed, editor)

    expect(controlTexts(dom)).toEqual(['☐'])
    expect(fields(xml, '<w:sdt>')).toBe(1)
    expect(dom.getElementsByTagName('w:p')[0].textContent).toBe('☐12')
  })

  it('when glyph and state disagree the state wins, and a tick makes them agree', async () => {
    const { parsed, editor } = await open(paragraph(`${controlBox(0, '☒')}${plain(' Label')}`))
    expect(editor.state.doc.child(0).textContent.startsWith('☐')).toBe(true)
    toggle(editor)
    const { dom } = await saved(parsed, editor)

    expect(controlTexts(dom)).toEqual(['☒'])
    expect(controlChecked(dom)).toBe('1')
  })
})

describe('clicking a line with a box', () => {
  /** the editor's click handling with the glyph drawn at x 10..20 */
  function click(editor: Editor, clientX: number) {
    editor.registerPlugin(boxTickPlugin)
    const view: EditorView = editor.view
    Range.prototype.getBoundingClientRect = vi.fn(
      () => ({ left: 10, right: 20, top: 100, bottom: 120 }) as DOMRect,
    )
    const event = new MouseEvent('click', { clientX, clientY: 110 })
    const pos = glyphPos(editor) + 1
    return view.someProp('handleClick', (handler) => handler(view, pos, event)) ?? false
  }
  const line = inCell(paragraph(`${legacyBox(1)}${formText('')}`))

  it('on the glyph ticks it and leaves the caret in the line', async () => {
    const { editor } = await open(line)
    const paragraphStart = editor.state.doc.resolve(glyphPos(editor)).start()
    expect(click(editor, 15)).toBe(true)
    expect(editor.state.doc.resolve(glyphPos(editor)).parent.textContent).toBe('☐')
    expect(editor.state.selection.empty).toBe(true)
    expect(editor.state.selection.from).toBe(paragraphStart + 1)
  })

  it('in the widened gap after the glyph leaves the tick alone', async () => {
    const { editor } = await open(line)
    expect(click(editor, 26)).toBe(false)
    expect(editor.state.doc.resolve(glyphPos(editor)).parent.textContent).toBe('☒')
  })

  it('beside the glyph leaves the tick and the caret to the click', async () => {
    const { editor } = await open(line)
    expect(click(editor, 200)).toBe(false)
    expect(editor.state.doc.resolve(glyphPos(editor)).parent.textContent).toBe('☒')
  })
})
