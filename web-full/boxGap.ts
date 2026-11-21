import type { StyleInfo } from '@genoffice/docx-engine'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Decoration, type EditorView } from '@tiptap/pm/view'

export const BOX_GAP_CLASS = 'vound-box-gap'

const BOX_MARKS = ['instrField', 'ctrlCheckbox']
const BOX_GLYPHS = ['☐', '☒', '☑']

/** an empty element right after the box: a caret after it sits outside the widened box, at the indent */
const gapEnd = (view: EditorView) => view.dom.ownerDocument.createElement('span')

/** whether the textblock starts with a form box */
export function opensWithBox(node: PmNode) {
  const first = node.firstChild
  return (
    first?.isText === true &&
    BOX_GLYPHS.includes(node.textContent[0]) &&
    first.marks.some((mark) => BOX_MARKS.includes(mark.type.name))
  )
}

/** Keeps a checkbox label inside the review highlight on hanging lines. */
export function boxGapDecorations(
  node: PmNode,
  pos: number,
  styles: ReadonlyMap<string, StyleInfo> | undefined,
) {
  if (node.type.name !== 'docParagraph' || !opensWithBox(node)) return []
  if (node.textContent[1] === '\t') return []
  const { styleId, indentFirstLine } = node.attrs
  const style = typeof styleId === 'string' ? styles?.get(styleId) : undefined
  const hangingTwips = -Number(indentFirstLine ?? style?.display?.indentFirstLineTwips ?? 0)
  if (hangingTwips <= 0) return []
  const glyph = pos + 1
  return [
    Decoration.inline(glyph, glyph + 1, {
      class: BOX_GAP_CLASS,
      style: `min-width:${hangingTwips / 20}pt`,
    }),
    Decoration.widget(glyph + 1, gapEnd, { side: -1, key: 'vound-box-gap' }),
  ]
}
