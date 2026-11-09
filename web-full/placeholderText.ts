// Where placeholder text lies: text carrying the editor's ctrlPlaceholder mark (a content control showing its
// placeholder, a form field's default text in a table cell).
import { EMPTY_CONTROL_TEXT } from '@genoffice/docx-engine'
import type { Node as PmNode } from '@tiptap/pm/model'

const PLACEHOLDER_MARK = 'ctrlPlaceholder'

type Range = { from: number; to: number }

const EN_SPACE = '\u2002'
/** the en spaces Word writes for an empty legacy text field */
const WORD_BLANK_LENGTH = 5

/** the offset in a plain-text paragraph where the en spaces of an empty form field result start: all of them
 *  in a paragraph of nothing else, or Word's run of them after a box or a label; null without any */
export function formFieldBlankStart(node: PmNode) {
  let plainText = node.childCount > 0
  node.forEach((child) => {
    plainText &&= child.isText
  })
  if (!plainText) return null
  const text = node.textContent
  let start = text.length
  while (start > 0 && text[start - 1] === EN_SPACE) start -= 1
  const length = text.length - start
  if (length === 0 || (start > 0 && length < WORD_BLANK_LENGTH)) return null
  return start
}

/** the zero-width filler an emptied inline control shows: a save and the host never see it */
export const isEmptyControlFiller = (node: PmNode) =>
  node.isText &&
  node.text === EMPTY_CONTROL_TEXT &&
  node.marks.some((mark) => mark.type.name === PLACEHOLDER_MARK)

/** the stretch of placeholder text in the paragraph at `pos` that reaches `pos`, edges included */
export function placeholderAt(doc: PmNode, pos: number): Range | null {
  const $pos = doc.resolve(pos)
  if (!$pos.parent.isTextblock) return null
  const ranges: Range[] = []
  $pos.parent.forEach((child, offset) => {
    if (!child.marks.some((mark) => mark.type.name === PLACEHOLDER_MARK)) return
    const from = $pos.start() + offset
    const last = ranges.at(-1)
    if (last?.to === from) last.to = from + child.nodeSize
    else ranges.push({ from, to: from + child.nodeSize })
  })
  return ranges.find((range) => range.from <= pos && pos <= range.to) ?? null
}
