/**
 * Typing over placeholder text (Vound fork): the typed text looks like the field's own text at once (no
 * placeholder mark or placeholder style), also when it arrives through a composition, and the save keeps the
 * content control around it with its properties, dropping only w:showingPlcHdr and the placeholder style.
 */
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { closeHistory } from '@tiptap/pm/history'
import { TextSelection } from '@tiptap/pm/state'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const STYLE =
  '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:rPr><w:color w:val="808080"/></w:rPr></w:style>'
const SDT_PR = (id: string) =>
  `<w:sdtPr><w:alias w:val="Rechtsform"/><w:tag w:val="form:${id}"/><w:id w:val="${id}"/><w:showingPlcHdr/><w:text/></w:sdtPr>`
const PH_RUN = (text: string) => `<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>${text}</w:t></w:r>`
const P_PR = '<w:pPr><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr>'
/** a cell whose only paragraph sits in a content control showing its placeholder */
const CELL_CONTROL = `<w:sdt>${SDT_PR('41')}<w:sdtContent><w:p>${P_PR}${PH_RUN('Rechtsform')}</w:p></w:sdtContent></w:sdt>`
const TABLE =
  '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>' +
  '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Rechtsform</w:t></w:r></w:p></w:tc>' +
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${CELL_CONTROL}</w:tc></w:tr></w:tbl>`
const BODY_CONTROL = `<w:sdt>${SDT_PR('42')}<w:sdtContent><w:p>${P_PR}${PH_RUN('Ort eingeben')}</w:p></w:sdtContent></w:sdt>`

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

async function open(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, extraStylesXml: STYLE }))
  const editor = new Editor({ element: document.createElement('div'), extensions: editorExtensions })
  editors.push(editor)
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  // the opened document is no step to undo
  editor.view.dispatch(closeHistory(editor.state.tr))
  return { parsed, editor }
}

/** where each stretch of placeholder text starts and ends */
function placeholders(editor: Editor) {
  const ranges: Array<{ from: number; to: number }> = []
  editor.state.doc.descendants((node, pos) => {
    if (node.isText && node.marks.some((mark) => mark.type.name === 'ctrlPlaceholder')) ranges.push({ from: pos, to: pos + node.nodeSize })
  })
  return ranges
}

/** the marks of the text node holding `text` */
function marksOf(editor: Editor, text: string) {
  let found: string[] = []
  editor.state.doc.descendants((node) => {
    if (node.isText && node.text === text)
      found = node.marks.map((mark) => (mark.type.name === 'docTextStyle' ? `docTextStyle:${mark.attrs.styleId ?? ''}` : mark.type.name))
  })
  return found
}

/** typing over a placeholder the way ProseMirror does it: its selection replaced by the text, marks across it */
function typeOver(editor: Editor, range: { from: number; to: number }, text: string) {
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, range.from, range.to)))
  editor.view.dispatch(editor.state.tr.insertText(text))
}

/** a composition read back from the DOM: the composed text lands in the placeholder's span, with its marks */
function composeOver(editor: Editor, range: { from: number; to: number }, text: string) {
  const marks = editor.state.doc.nodeAt(range.from)!.marks
  editor.view.dispatch(editor.state.tr.replaceWith(range.from, range.to, editor.schema.text(text, marks)).setMeta('composition', 1))
}

async function saved(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const out = await saveDocx(parsed, pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks)
  return { out, xml: await (await JSZip.loadAsync(out)).file('word/document.xml')!.async('string') }
}

describe('typing over placeholder text', () => {
  it('gives the typed text the field formatting without the placeholder mark or style', async () => {
    const { editor } = await open(`<w:p><w:sdt>${SDT_PR('40')}<w:sdtContent>${PH_RUN('Name eingeben')}</w:sdtContent></w:sdt></w:p>`)
    typeOver(editor, placeholders(editor)[0], 'Muster GmbH')
    expect(marksOf(editor, 'Muster GmbH')).not.toContain('ctrlPlaceholder')
    expect(marksOf(editor, 'Muster GmbH').filter((m) => m.startsWith('docTextStyle:Placeholder'))).toEqual([])
  })

  it('does the same for text a composition (dead keys, umlauts) puts there', async () => {
    const { editor } = await open(`<w:p><w:sdt>${SDT_PR('43')}<w:sdtContent>${PH_RUN('Ort eingeben')}</w:sdtContent></w:sdt></w:p>`)
    composeOver(editor, placeholders(editor)[0], 'Lüneburg')
    expect(marksOf(editor, 'Lüneburg')).toEqual([])
  })

  it('brings the placeholder back whole on undo', async () => {
    const { editor } = await open(`<w:p><w:sdt>${SDT_PR('44')}<w:sdtContent>${PH_RUN('Ort eingeben')}</w:sdtContent></w:sdt></w:p>`)
    typeOver(editor, placeholders(editor)[0], 'Leipzig')
    editor.commands.undo()
    expect(editor.state.doc.textContent).toBe('Ort eingeben')
    expect(marksOf(editor, 'Ort eingeben')).toEqual(['ctrlPlaceholder', 'docTextStyle:PlaceholderText'])
  })

  it('keeps a content control around a cell paragraph on save, without its placeholder flag and style', async () => {
    const { parsed, editor } = await open(TABLE)
    typeOver(editor, placeholders(editor)[0], 'Muster GmbH')
    const { out, xml } = await saved(editor, parsed)
    const expected = TABLE.replace('<w:showingPlcHdr/>', '').replace(PH_RUN('Rechtsform'), '<w:r><w:t xml:space="preserve">Muster GmbH</w:t></w:r>')
    expect(xml).toContain(expected)
    const reparsed = await parseDocx(out)
    const runs = reparsed.blocks[0].table!.rows[0][1].richParas![0].runs
    expect(runs.map((run) => [run.text, run.placeholder === true, run.styleId ?? null])).toEqual([['Muster GmbH', false, null]])
  })

  it('keeps a content control around a body paragraph on save, without its placeholder flag and style', async () => {
    const { parsed, editor } = await open(`<w:p><w:r><w:t>Ort:</w:t></w:r></w:p>${BODY_CONTROL}`)
    typeOver(editor, placeholders(editor)[0], 'Leipzig')
    const { xml } = await saved(editor, parsed)
    expect(xml).toContain(BODY_CONTROL.replace('<w:showingPlcHdr/>', '').replace(PH_RUN('Ort eingeben'), '<w:r><w:t xml:space="preserve">Leipzig</w:t></w:r>'))
  })
})
