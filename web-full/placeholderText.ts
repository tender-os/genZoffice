// Where placeholder text lies: text carrying the editor's ctrlPlaceholder mark (a content control showing its
// placeholder, a form field's default text in a table cell).
import type { Node as PmNode } from '@tiptap/pm/model'

const PLACEHOLDER_MARK = 'ctrlPlaceholder'

type Range = { from: number; to: number }

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
