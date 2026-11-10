// Where a restyle goes when the field's text names no single place (fieldStyle.ts): the legacy text form field
// the host's index names, or, for a field with no text yet, the block's one empty field, else the end of its
// text, so the style is stored for the text typed there next; of several empty fields, the one at the caret. An emptied field's filler takes no mark (the editor
// drops a filler an edit touched), so its place is the caret at its start.
import type { Node as PmNode } from '@tiptap/pm/model'
import type { FieldTextStyle } from './embed-protocol'
import type { FieldTarget } from './fieldWrite'
import { locateBlock } from './markers'
import { isEmptyControlFiller } from './placeholderText'

type Paragraph = { node: PmNode; pos: number }
type Range = { kind: 'text'; from: number; to: number }

const FIELD_MARK = 'formTextField'
const EN_SPACE = ' '

const fieldIndexOf = (node: PmNode) => node.marks.find((mark) => mark.type.name === FIELD_MARK)?.attrs.index

function paragraphsOf(block: Paragraph): Paragraph[] {
  if (block.node.isTextblock) return [block]
  const paragraphs: Paragraph[] = []
  block.node.descendants((node, offset) => {
    if (node.isTextblock) paragraphs.push({ node, pos: block.pos + 1 + offset })
    return !node.isTextblock
  })
  return paragraphs
}

/** each text node of the paragraphs with its document range */
const textsOf = (paragraphs: Paragraph[]) =>
  paragraphs.flatMap(({ node, pos }) => {
    const texts: { node: PmNode; from: number; to: number }[] = []
    node.forEach((child, offset) => {
      if (child.isText) texts.push({ node: child, from: pos + 1 + offset, to: pos + 1 + offset + child.nodeSize })
    })
    return texts
  })

type Text = ReturnType<typeof textsOf>[number]

/** the place of these texts of one field: all of them, or the caret before an emptied field's filler */
function placeOfTexts(texts: Text[]): Range {
  const from = texts[0].from
  return texts.every((text) => isEmptyControlFiller(text.node)) ? { kind: 'text', from, to: from } : { kind: 'text', from, to: texts[texts.length - 1].to }
}

/** the text of the form field with this index, when one paragraph of the block holds it */
function formFieldRange(paragraphs: Paragraph[], index: number): Range | null {
  const holding = paragraphs.filter(({ node }) => textsOf([{ node, pos: 0 }]).some((text) => fieldIndexOf(text.node) === index))
  if (holding.length !== 1) return null
  return placeOfTexts(textsOf(holding).filter((text) => fieldIndexOf(text.node) === index))
}

/** an emptied control's or form field's filler, or a blank form field of en spaces */
const isEmptyField = (node: PmNode) =>
  isEmptyControlFiller(node) || (fieldIndexOf(node) !== undefined && node.text === EN_SPACE.repeat(node.text?.length ?? 0))

function emptyFieldRange(paragraphs: Paragraph[], caret: number): Range | null {
  const empties = textsOf(paragraphs).filter((text) => isEmptyField(text.node))
  const atCaret = empties.filter((text) => text.from <= caret && caret <= text.to)
  if (empties.length === 1) return placeOfTexts(empties)
  if (atCaret.length === 1) return placeOfTexts(atCaret)
  const last = paragraphs.at(-1)
  if (empties.length > 1 || !last) return null
  const end = last.pos + last.node.nodeSize - 1
  return { kind: 'text', from: end, to: end }
}

/** `caret` picks among several empty fields: the one the reviewer is typing in */
export function emptyFieldPlace(
  doc: PmNode,
  request: FieldTarget & { formText?: number; style: FieldTextStyle },
  caret: number,
): Range | null {
  if (request.replaces === 'gap') return null
  const block = locateBlock(doc, request.block)
  if (!block) return null
  const paragraphs = paragraphsOf(block)
  if (request.formText !== undefined) return formFieldRange(paragraphs, request.formText)
  return request.previous === '' ? emptyFieldRange(paragraphs, caret) : null
}
