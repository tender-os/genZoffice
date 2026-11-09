import { isOn } from './checkbox-control'
import type { ControlEnter } from './types'
import { attrsOf, childrenOf, findChild, nameOf, xmlParser, type XNode } from './xml-utils'

/** w:sdtPr children that make a control something other than rich text */
const TYPED_CONTROL = new Set([
  'w:text',
  'w:date',
  'w:dropDownList',
  'w:comboBox',
  'w:picture',
  'w:docPartObj',
  'w:docPartList',
  'w:citation',
  'w:bibliography',
  'w:equation',
  'w:group',
  'w14:checkbox',
  'w15:repeatingSection',
  'w15:repeatingSectionItem',
])

/** paragraph children that hold no content of their own */
const MARKUP_ONLY = new Set([
  'w:pPr',
  'w:bookmarkStart',
  'w:bookmarkEnd',
  'w:proofErr',
  'w:commentRangeStart',
  'w:commentRangeEnd',
  'w:permStart',
  'w:permEnd',
])

/** what Enter does in a block-level content control (its w:sdt node) */
export function controlEnterOf(sdt: XNode): ControlEnter {
  const sdtPr = findChild(sdt, 'w:sdtPr')
  const kinds = childrenOf(sdtPr ?? {}).flatMap((child) => {
    const name = nameOf(child)
    return name && TYPED_CONTROL.has(name) ? [child] : []
  })
  if (kinds.length === 0) return 'paragraph'
  // an absent w:multiLine is off, unlike an on/off element without a value
  const multiLine = attrsOf(kinds[0])['w:multiLine']
  const lines =
    kinds.length === 1 &&
    nameOf(kinds[0]) === 'w:text' &&
    multiLine !== undefined &&
    isOn(multiLine)
  return lines ? 'lineBreak' : 'none'
}

/** what Enter does in a w:p holding inline content controls: a line break in a rich-text or multi-line
 *  control that fills the paragraph, nothing otherwise */
export function inlineControlEnterOf(p: XNode): ControlEnter {
  const content = childrenOf(p).filter((child) => {
    const name = nameOf(child)
    return name !== undefined && !MARKUP_ONLY.has(name)
  })
  if (content.length !== 1 || nameOf(content[0]) !== 'w:sdt') return 'none'
  return controlEnterOf(content[0]) === 'none' ? 'none' : 'lineBreak'
}

type WrappingControl = { enter: ControlEnter; control: number }

/** the paragraphs under `node` (a table cell) that block-level controls wrap, each with what Enter does in the
 *  innermost one and that control's number within `node` */
export function blockControlParagraphs(node: XNode): Map<XNode, WrappingControl> {
  const found = new Map<XNode, WrappingControl>()
  let count = 0
  const visit = (parent: XNode, wrapping: WrappingControl | null) => {
    for (const child of childrenOf(parent)) {
      const name = nameOf(child)
      if (name === 'w:p' && wrapping) found.set(child, wrapping)
      if (name !== 'w:sdt') continue
      const content = findChild(child, 'w:sdtContent')
      if (content) visit(content, { enter: controlEnterOf(child), control: count++ })
    }
  }
  visit(node, null)
  return found
}

/** what Enter does in a block-level content control, from its XML */
export function controlEnterOfXml(sdtXml: string): ControlEnter {
  const nodes: XNode[] = xmlParser.parse(sdtXml)
  const sdt = nodes.find((node) => nameOf(node) === 'w:sdt')
  return sdt ? controlEnterOf(sdt) : 'none'
}
