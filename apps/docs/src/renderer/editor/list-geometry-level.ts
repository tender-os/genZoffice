import { levelWithStyleIndent, type NumberingLevel, type StyleInfo } from '@genoffice/docx-engine'
import type { Node as PmNode } from '@tiptap/pm/model'

/** The level whose w:ind lays out a list item: one numbered through its style keeps that style's indent. */
export function geometryLevelOf(
  level: NumberingLevel,
  node: PmNode,
  styles: Map<string, StyleInfo> | undefined,
): NumberingLevel {
  if (!node.attrs.listFromStyle) return level
  return levelWithStyleIndent(level, styles?.get(node.attrs.styleId)?.listIndent)
}
