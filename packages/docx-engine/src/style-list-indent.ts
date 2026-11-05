import type { NumberingLevel, StyleInfo } from './types'
import { findChild, type XNode } from './xml-utils'

type ListIndent = NonNullable<StyleInfo['listIndent']>

/** w:ind components a style chain declares from the style that sets the numId down. */
export function styleListIndent(
  styleId: string,
  ownIndents: Map<string, ListIndent>,
  numIdStyles: Set<string>,
  basedOn: Map<string, string>,
): ListIndent {
  let indent: ListIndent = {}
  const seen = new Set<string>()
  for (let id: string | undefined = styleId; id && !seen.has(id); id = basedOn.get(id)) {
    seen.add(id)
    indent = { ...ownIndents.get(id), ...indent }
    if (numIdStyles.has(id)) break
  }
  return indent
}

/** No w:numId of the paragraph's own: its list number comes from the paragraph style. */
export function numIdFromStyle(pPr: XNode | undefined): boolean {
  const numPr = pPr ? findChild(pPr, 'w:numPr') : undefined
  return !numPr || !findChild(numPr, 'w:numId')
}

/** The level with a style's list indent laid over its w:ind. */
export function levelWithStyleIndent(
  level: NumberingLevel,
  indent: ListIndent | undefined,
): NumberingLevel {
  if (!indent) return level
  const out = { ...level }
  if (indent.indentLeftTwips !== undefined) out.indentLeft = indent.indentLeftTwips
  if (indent.indentRightTwips !== undefined) out.indentRight = indent.indentRightTwips
  const firstLine = indent.indentFirstLineTwips
  if (firstLine !== undefined) {
    out.hanging = firstLine < 0 ? -firstLine : undefined
    out.firstLine = firstLine > 0 ? firstLine : undefined
  }
  return out
}
