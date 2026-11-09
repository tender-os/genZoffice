/** Every paragraph-level node carries its source paragraph index (docxPara,
 * rendered as data-src-p): top-level paragraphs, cell paragraphs and protected
 * blocks. A paragraph the user inserts has none. */
import { Editor, type JSONContent } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { splitGridCell } from '../src/renderer/editor/table-ops'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const P = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
const TC = (inner: string, tcPr = '') => `<w:tc>${tcPr}${inner}</w:tc>`
const FORM_FIELD_P =
  '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>entry</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
const LIST_P =
  '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
  '<w:r><w:t>cell item</w:t></w:r></w:p>'
const BODY =
  '<w:p/>' +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${TC(P('r0c0'))}${TC(P('r0c1'), '<w:tcPr><w:vMerge w:val="restart"/></w:tcPr>')}</w:tr>` +
  `<w:tr>${TC(P('r1c0a') + LIST_P)}${TC('<w:p/>', '<w:tcPr><w:vMerge/></w:tcPr>')}</w:tr>` +
  '</w:tbl>' +
  FORM_FIELD_P +
  P('last')

const PARAGRAPH_LEVEL = new Set(['docParagraph', 'docHeading', 'docListItem', 'docProtected'])

function paragraphLevelNodes(node: JSONContent, out: JSONContent[] = []): JSONContent[] {
  if (node.type && PARAGRAPH_LEVEL.has(node.type)) out.push(node)
  for (const child of node.content ?? []) paragraphLevelNodes(child, out)
  return out
}

async function open(bytes?: Uint8Array) {
  const parsed = await parseDocx(bytes ?? (await buildDocx({ bodyXml: BODY, withNumbering: true })))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { parsed, editor }
}

describe('docxPara: the source paragraph index on editor nodes', () => {
  it('equals sourceParaIndex on top-level, cell and protected paragraphs', async () => {
    const { parsed, editor } = await open()
    const blocks = parsed.blocks.filter((b) => !b.hidden)
    const rows = blocks[1].table!.rows
    const nodes = paragraphLevelNodes(editor.getJSON())
    expect(nodes.map((n) => [n.type, n.attrs?.docxPara])).toEqual([
      ['docParagraph', blocks[0].sourceParaIndex],
      ['docParagraph', rows[0][0].richParas![0].sourceParaIndex],
      ['docParagraph', rows[0][1].richParas![0].sourceParaIndex],
      ['docParagraph', rows[1][0].richParas![0].sourceParaIndex],
      ['docListItem', rows[1][0].richParas![1].sourceParaIndex],
      // the vMerge continuation cell is not rendered
      ['docProtected', blocks[2].sourceParaIndex],
      ['docParagraph', blocks[3].sourceParaIndex],
    ])
    expect(nodes.map((n) => n.attrs?.docxPara)).toEqual([0, 1, 2, 3, 4, 6, 7])
    editor.destroy()
  })

  it('renders data-src-p on paragraphs and protected blocks', async () => {
    const { editor } = await open()
    const dom = editor.view.dom
    const marked = [...dom.querySelectorAll('[data-src-p]')].map((el) =>
      el.getAttribute('data-src-p'),
    )
    expect(marked).toEqual(['0', '1', '2', '3', '4', '6', '7'])
    expect(dom.querySelector('[data-doc-protected][data-src-p="6"]')).not.toBeNull()
    editor.destroy()
  })

  it('is null on a paragraph the user inserts', async () => {
    const { editor } = await open()
    editor.commands.insertContentAt(0, {
      type: 'docParagraph',
      content: [{ type: 'text', text: 'new' }],
    })
    const first = paragraphLevelNodes(editor.getJSON())[0]
    expect(first.content?.[0]?.text).toBe('new')
    expect(first.attrs?.docxPara).toBeNull()
    expect(editor.view.dom.querySelector('p')?.hasAttribute('data-src-p')).toBe(false)
    editor.destroy()
  })
})

/** position and node of the first textblock whose text is `text` */
function textblockWithText(editor: Editor, text: string): { node: ProseMirrorNode; pos: number } {
  let found: { node: ProseMirrorNode; pos: number } | null = null
  editor.state.doc.descendants((node, pos) => {
    if (found) return false
    if (node.isTextblock && node.textContent === text) found = { node, pos }
    return !found
  })
  if (!found) throw new Error(`no paragraph with text ${text}`)
  return found
}

function nodesWithDocxPara(editor: Editor, index: number): string[] {
  const texts: string[] = []
  editor.state.doc.descendants((node) => {
    if (node.attrs.docxPara === index) texts.push(node.textContent)
  })
  return texts
}

describe('docxPara on paragraphs the editor makes from stamped ones', () => {
  it('a split keeps the index on the first half and makes the second its continuation', async () => {
    const { editor } = await open()
    const { pos } = textblockWithText(editor, 'last')
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos + 1 + 2)),
    )
    expect(editor.commands.splitBlock()).toBe(true)
    expect(nodesWithDocxPara(editor, 7)).toEqual(['la'])
    expect(textblockWithText(editor, 'st').node.attrs.continuesPara).toBe(7)
    editor.destroy()
  })

  it('a blank cell paragraph made from a stamped template cell carries null', async () => {
    const { editor } = await open()
    let table: ProseMirrorNode | null = null
    editor.state.doc.forEach((node) => {
      if (!table && node.type.name === 'docTable') table = node
    })
    const template = table!.child(0).child(0).firstChild!
    expect(template.attrs.docxPara).toBe(1)
    const result = splitGridCell(table!, { row: 0, col: 0 }, 1, 2)!
    const blank = result.table.child(0).child(1).firstChild!
    expect(blank.type.name).toBe('docParagraph')
    expect(blank.textContent).toBe('')
    expect(blank.attrs.docxIndex).toBeNull()
    expect(blank.attrs.docxPara).toBeNull()
    editor.destroy()
  })
})

describe('saving an edited document through the editor', () => {
  const documentXmlOf = async (bytes: Uint8Array) =>
    (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
  const COUNTED = [/<w:highlight[\s>/]/g, /<w:shd[\s>/]/g, /<w:bookmarkStart[\s>/]/g]

  for (const target of ['last', 'r0c0']) {
    it(`writes no marker name and no new decoration after editing "${target}"`, async () => {
      const bytes = await buildDocx({ bodyXml: BODY, withNumbering: true })
      const { parsed, editor } = await open(bytes)
      const { node, pos } = textblockWithText(editor, target)
      expect(node.attrs.docxPara).toEqual(expect.any(Number))
      editor.view.dispatch(
        editor.state.tr
          .setSelection(TextSelection.create(editor.state.doc, pos + 1))
          .insertText('edited '),
      )
      const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
      expect(plan.changedCount).toBeGreaterThan(0)
      const saved = await saveDocx(parsed, plan.saveBlocks)
      editor.destroy()

      const sourceXml = await documentXmlOf(bytes)
      const savedXml = await documentXmlOf(saved)
      expect(savedXml).toContain(`edited ${target}`)
      for (const name of ['sourceParaIndex', 'docxPara', 'data-src-p']) {
        expect(savedXml).not.toContain(name)
      }
      for (const pattern of COUNTED) {
        expect(savedXml.match(pattern)?.length ?? 0).toBe(sourceXml.match(pattern)?.length ?? 0)
      }
      expect((await JSZip.loadAsync(saved)).file('word/comments.xml')).toBeNull()
    })
  }
})
