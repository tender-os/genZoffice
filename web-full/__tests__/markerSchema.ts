// A minimal stand-in for the Docs schema: the node names the markers code looks for, docxPara, the
// run-formatting mark (docTextStyle) and line break (docBreak) the review-mode paste tests read, and the
// field passthrough block (docProtected with fieldDisplay) the field writes edit, and placeholder text
// (ctrlPlaceholder) the review mode selects whole.
import { Schema, type Node as PmNode } from '@tiptap/pm/model'

export const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    docParagraph: {
      group: 'block',
      content: 'inline*',
      attrs: { docxPara: { default: null }, align: { default: null } },
      // only the view tests (scrollToMarker, reviewMode, fieldWrite, fieldReporting) render or parse
      toDOM: () => ['p', 0],
      parseDOM: [{ tag: 'p' }],
    },
    docProtected: {
      group: 'block',
      atom: true,
      attrs: { docxPara: { default: null }, fieldDisplay: { default: null }, styleId: { default: null } },
      toDOM: () => ['div', { class: 'doc-protected' }],
    },
    docTable: { group: 'block', content: 'docTableRow+', toDOM: () => ['table', ['tbody', 0]] },
    docTableRow: { content: 'docTableCell+', toDOM: () => ['tr', 0] },
    docTableCell: { content: 'docParagraph+', toDOM: () => ['td', 0] },
    docBreak: { inline: true, group: 'inline', toDOM: () => ['br'], parseDOM: [{ tag: 'br' }] },
    text: { group: 'inline' },
  },
  marks: {
    ctrlPlaceholder: { inclusive: false, toDOM: () => ['span', { 'data-ctrl-placeholder': '' }, 0] },
    instrField: { inclusive: false, attrs: { instr: { default: '' } }, toDOM: () => ['span', 0] },
    docTextStyle: {
      attrs: { bold: { default: null }, font: { default: null } },
      toDOM: (mark) => ['span', { 'data-bold': mark.attrs.bold, 'data-font': mark.attrs.font }, 0],
      parseDOM: [{ tag: 'b', attrs: { bold: true } }],
    },
  },
})

export const para = (docxPara: number | null, text = '') =>
  schema.nodes.docParagraph.create({ docxPara }, text ? schema.text(text) : null)
/** a paragraph holding `before`, then placeholder text (the ctrlPlaceholder mark), then `after` */
export const placeholderPara = (docxPara: number, before: string, placeholder: string, after = '') =>
  schema.nodes.docParagraph.create({ docxPara }, [
    ...(before ? [schema.text(before)] : []),
    schema.text(placeholder, [schema.marks.ctrlPlaceholder.create()]),
    ...(after ? [schema.text(after)] : []),
  ])
export const protectedBlock = (docxPara: number) => schema.nodes.docProtected.create({ docxPara })
/** a legacy form-field paragraph as the editor parses it: a passthrough block whose visible text is
 *  fieldDisplay.left */
export const fieldPara = (docxPara: number, left: string, extra: Record<string, unknown> = {}) =>
  schema.nodes.docProtected.create({ docxPara, styleId: 'Formular', fieldDisplay: { kind: 'text', left, fontFamily: 'Arial', ...extra } })
export const cell = (...paragraphs: PmNode[]) => schema.nodes.docTableCell.create(null, paragraphs)
export const table = (...cells: PmNode[]) => schema.nodes.docTable.create(null, schema.nodes.docTableRow.create(null, cells))
export const doc = (...blocks: PmNode[]) => schema.nodes.doc.create(null, blocks)
/** paragraphs with source indices from..to-1, each with some text */
export const paragraphs = (from: number, to: number) =>
  Array.from({ length: to - from }, (_, i) => para(from + i, `paragraph ${from + i}`))
