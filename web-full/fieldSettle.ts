// The host's vound:settleField and vound:reopenField: a field the reviewer accepted or left blank is no longer a
// place to fill, so the author's fill-here highlight goes, on screen and in the saved file; reopening the field
// brings it back. Each is a host write in its own undo step.
import type { Node as PmNode } from '@tiptap/pm/model'
import type { EditorView } from '@tiptap/pm/view'
import type { BlockRef } from './embed-protocol'
import { KEEPS_PLACEHOLDERS_META } from '../apps/docs/src/renderer/editor/placeholder-typing'
import { HOST_WRITE_META } from './fieldWrite'
import { reopenFormTextFields, settleFormTextFields } from './formTextField'
import { locateBlock } from './markers'

/** what settling or reopening does to a field passthrough's display, which keeps its own text */
const PASSTHROUGH = {
  settle: (display: Record<string, unknown>) => ({ ...display, settled: true }),
  reopen: ({ settled: _settled, ...display }: Record<string, unknown>) => display,
}
const TEXT_FIELDS = { settle: settleFormTextFields, reopen: reopenFormTextFields }
type FieldChange = keyof typeof TEXT_FIELDS

/** the passthroughs of legacy text fields in a block, and the block itself when it is one */
function textPassthroughs(node: PmNode, pos: number) {
  const found: { node: PmNode; pos: number }[] = []
  const visit = (child: PmNode, at: number) => {
    if (child.attrs.fieldDisplay?.kind === 'text') found.push({ node: child, pos: at })
  }
  visit(node, pos)
  node.descendants((child, offset) => visit(child, pos + 1 + offset))
  return found
}

function changeFields(view: EditorView, block: BlockRef, change: FieldChange) {
  const located = locateBlock(view.state.doc, block)
  if (!located) return
  const { doc, tr } = view.state
  TEXT_FIELDS[change](doc, located.pos, located.pos + located.node.nodeSize, tr)
  for (const { node, pos } of textPassthroughs(located.node, located.pos)) {
    const display = PASSTHROUGH[change](node.attrs.fieldDisplay)
    if (display.settled !== node.attrs.fieldDisplay.settled)
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, fieldDisplay: display })
  }
  if (tr.docChanged)
    view.dispatch(tr.setMeta(HOST_WRITE_META, true).setMeta(KEEPS_PLACEHOLDERS_META, true))
}

/** settles the legacy text fields of one marked block; nothing when the block is gone or they already are */
export const settleField = (view: EditorView, block: BlockRef) =>
  changeFields(view, block, 'settle')

/** reopens the settled legacy text fields of one marked block */
export const reopenField = (view: EditorView, block: BlockRef) =>
  changeFields(view, block, 'reopen')
