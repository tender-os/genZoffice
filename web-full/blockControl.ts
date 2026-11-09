// The block-level content control a paragraph sits in, as the editor's paragraph attributes say.
import type { Node as PmNode } from '@tiptap/pm/model'

/** a key for the block-level control spanning several paragraphs that holds this one, or null: its number in
 *  the paragraph's table cell (cellControl), or its group among the body's blocks (sdtShell) */
export function blockControlKey(node: PmNode) {
  if (typeof node.attrs.cellControl === 'number') return `cell ${node.attrs.cellControl}`
  if (typeof node.attrs.sdtShell !== 'string') return null
  const shell: { group?: number } = JSON.parse(node.attrs.sdtShell)
  return shell.group === undefined ? null : `group ${shell.group}`
}
