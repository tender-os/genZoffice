// Vound fork: the vertical alignment of the cells and the alignment of their paragraphs, as the reviewer set them
// (the host's Format control), written into a table kept as its original XML because a rebuild would lose its form
// fields or content controls (regeneratedTableXml).
import {
  withCellVerticalAlignment,
  withParagraphAlignment,
  xmlSegments,
  type Block,
  type TableCell,
  type TableModel,
} from '@genoffice/docx-engine'

/** the top-level children `tag` of an element, with their range in it, its own tags left out */
function inner(xml: string, tag: string) {
  const open = xml.indexOf('>') + 1
  const close = xml.lastIndexOf('</')
  return xmlSegments(xml, tag, open, close)
}

function withCellFormat(cellXml: string, current: TableCell, original: TableCell) {
  const paragraphs = inner(cellXml, 'w:p')
  const before = original.richParas ?? []
  const after = current.richParas ?? []
  let out = cellXml
  if (paragraphs.length === after.length && after.length === before.length) {
    for (let i = paragraphs.length - 1; i >= 0; i--) {
      const align = after[i].align
      if (!align || align === before[i].align) continue
      const { start, end } = paragraphs[i]
      out =
        out.slice(0, start) + withParagraphAlignment(out.slice(start, end), align) + out.slice(end)
    }
  }
  return current.vAlign && current.vAlign !== original.vAlign
    ? withCellVerticalAlignment(out, current.vAlign)
    : out
}

const realCells = (row: TableCell[] | undefined) => (row ?? []).filter((cell) => !cell.gridGap)

/** the table XML with each cell's and paragraph's alignment the model changed */
export function withCellFormats(tableXml: string, model: TableModel, original: Block) {
  const rows = original.table?.rows
  if (!rows) return tableXml
  let out = tableXml
  const rowSegments = inner(out, 'w:tr')
  for (let r = rowSegments.length - 1; r >= 0; r--) {
    const row = rowSegments[r]
    const rowXml = out.slice(row.start, row.end)
    const cells = inner(rowXml, 'w:tc')
    const currentCells = realCells(model.rows[r])
    const originalCells = realCells(rows[r])
    if (cells.length !== currentCells.length || cells.length !== originalCells.length) continue
    let next = rowXml
    for (let c = cells.length - 1; c >= 0; c--) {
      const { start, end } = cells[c]
      next =
        next.slice(0, start) +
        withCellFormat(next.slice(start, end), currentCells[c], originalCells[c]) +
        next.slice(end)
    }
    out = out.slice(0, row.start) + next + out.slice(row.end)
  }
  return out
}
