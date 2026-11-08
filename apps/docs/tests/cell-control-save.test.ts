/**
 * A content control around paragraphs of a table cell (Vound fork) keeps its wrapper through edits that change
 * the cell's paragraph count: a paragraph made inside the control (Enter, a paste) stays in it, paragraphs
 * removed from it leave the wrapper around the rest, and a control whose paragraphs all went keeps one empty
 * paragraph. Everything else in the cell keeps its bytes.
 */
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { Fragment, Slice } from '@tiptap/pm/model'
import { TextSelection } from '@tiptap/pm/state'
import { afterEach, describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const P_PR = '<w:pPr><w:jc w:val="center"/></w:pPr>'
const para = (text: string) => `<w:p>${P_PR}<w:r><w:t>${text}</w:t></w:r></w:p>`
const SDT_OPEN = '<w:sdt><w:sdtPr><w:alias w:val="Anschrift"/><w:tag w:val="form:7"/><w:id w:val="7"/><w:richText/></w:sdtPr><w:sdtContent>'
const SDT_CLOSE = '</w:sdtContent></w:sdt>'
const BEFORE = '<w:p><w:r><w:t>Vorher</w:t></w:r></w:p>'
const AFTER = '<w:p><w:r><w:t>Hinweis</w:t></w:r></w:p>'
const TC_OPEN = '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
const cellXml = (inside: string) => `${TC_OPEN}${BEFORE}${SDT_OPEN}${inside}${SDT_CLOSE}${AFTER}</w:tc>`
const TABLE = `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${cellXml(para('Zeile eins') + para('Zeile zwei'))}</w:tr></w:tbl><w:p/>`

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

async function open() {
  const parsed = await parseDocx(await buildDocx({ bodyXml: TABLE }))
  const editor = new Editor({ element: document.createElement('div'), extensions: editorExtensions })
  editors.push(editor)
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  return { parsed, editor }
}

/** the end of the cell paragraph with this text */
function endOf(editor: Editor, text: string) {
  let end = -1
  editor.state.doc.descendants((node, pos) => {
    if (node.isTextblock && node.textContent === text) end = pos + node.nodeSize - 1
  })
  return end
}

async function savedCell(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const out = await saveDocx(parsed, pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks)
  const xml = await (await JSZip.loadAsync(out)).file('word/document.xml')!.async('string')
  const reparsed = await parseDocx(out)
  const texts = reparsed.blocks[0].table!.rows[0][0].richParas!.map((p) => p.runs.map((run) => run.text).join(''))
  return { cell: xml.slice(xml.indexOf(TC_OPEN), xml.indexOf('</w:tc>') + '</w:tc>'.length), texts }
}

const emitted = (text: string) => `<w:p>${P_PR}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`

describe('a content control around cell paragraphs on save', () => {
  it('keeps a paragraph made by Enter inside the control', async () => {
    const { parsed, editor } = await open()
    editor.commands.setTextSelection(endOf(editor, 'Zeile eins'))
    editor.commands.keyboardShortcut('Enter')
    editor.commands.insertContent('neu')
    const { cell, texts } = await savedCell(editor, parsed)
    // the paragraph Enter split is written from its runs, the other one keeps its bytes
    expect(cell).toBe(cellXml(emitted('Zeile eins') + emitted('neu') + para('Zeile zwei')))
    expect(texts).toEqual(['Vorher', 'Zeile eins', 'neu', 'Zeile zwei', 'Hinweis'])
  })

  it('keeps pasted paragraphs inside the control they were pasted into', async () => {
    const { parsed, editor } = await open()
    const at = endOf(editor, 'Zeile zwei')
    const p = (text: string) => editor.schema.nodes.docParagraph.create(null, text ? editor.schema.text(text) : null)
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at)).replaceSelection(new Slice(Fragment.from([p(''), p('A'), p('B')]), 1, 1)))
    const { cell, texts } = await savedCell(editor, parsed)
    expect(texts).toEqual(['Vorher', 'Zeile eins', 'Zeile zwei', 'A', 'B', 'Hinweis'])
    expect(cell.startsWith(`${TC_OPEN}${BEFORE}${SDT_OPEN}${para('Zeile eins')}`)).toBe(true)
    expect(cell.endsWith(`${SDT_CLOSE}${AFTER}</w:tc>`)).toBe(true)
  })

  it('keeps the wrapper around what is left when a paragraph inside it goes', async () => {
    const { parsed, editor } = await open()
    editor.commands.deleteRange({ from: endOf(editor, 'Zeile eins') - 'Zeile eins'.length, to: endOf(editor, 'Zeile zwei') - 'Zeile zwei'.length })
    const { cell, texts } = await savedCell(editor, parsed)
    expect(texts).toEqual(['Vorher', 'Zeile zwei', 'Hinweis'])
    expect(cell).toBe(cellXml(emitted('Zeile zwei')))
  })

  it('keeps one empty paragraph in a control whose paragraphs all went', async () => {
    const { parsed, editor } = await open()
    editor.commands.deleteRange({ from: endOf(editor, 'Vorher'), to: endOf(editor, 'Zeile zwei') })
    const { cell, texts } = await savedCell(editor, parsed)
    expect(texts).toEqual(['Vorher', '', 'Hinweis'])
    expect(cell).toBe(cellXml(`<w:p>${P_PR}</w:p>`))
  })
})
