// A box at the start of a hanging line with nothing separating it from what follows: Word sets the text right after
// the box, left of the indent, so it starts outside the field's highlight, which begins at the indent. In review mode
// the box is drawn as wide as the hanging indent, so the caret and the text start at the indent. Display only: these
// are decorations, nothing reaches the document or a save. A line whose box is followed by a tab or a space
// keeps Word's layout: its tab already reaches the indent, and its label is the form's own.
import type { Node as PmNode } from '@tiptap/pm/model'
import { Decoration, type EditorView } from '@tiptap/pm/view'

export const BOX_GAP_CLASS = 'vound-box-gap'

const BOX_MARKS = ['instrField', 'ctrlCheckbox']
const BOX_GLYPHS = ['☐', '☒', '☑']
const SEPARATORS = ['\t', ' ', ' ']

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

/** the decorations that start the text of the paragraph at `pos` at its hanging indent; none for any other */
export function boxGapDecorations(node: PmNode, pos: number) {
  const hangingTwips = -Number(node.attrs.indentFirstLine ?? 0)
  if (node.type.name !== 'docParagraph' || hangingTwips <= 0 || !opensWithBox(node)) return []
  if (SEPARATORS.includes(node.textContent[1])) return []
  const glyph = pos + 1
  return [
    Decoration.inline(glyph, glyph + 1, {
      class: BOX_GAP_CLASS,
      style: `min-width:${hangingTwips / 20}pt`,
    }),
    Decoration.widget(glyph + 1, gapEnd, { side: -1, key: 'vound-box-gap' }),
  ]
}
