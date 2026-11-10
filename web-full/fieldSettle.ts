// The host's vound:settleField: a field the reviewer accepted or left blank is no longer a place to fill, so the
// author's fill-here highlight goes, on screen and in the saved file. A host write in its own undo step.
import type { EditorView } from '@tiptap/pm/view'
import type { BlockRef } from './embed-protocol'
import { HOST_WRITE_META } from './fieldWrite'
import { settleFormTextFields } from './formTextField'
import { locateBlock } from './markers'

/** settles the legacy text fields of one marked block; nothing when the block is gone or they already are */
export function settleField(view: EditorView, block: BlockRef) {
  const located = locateBlock(view.state.doc, block)
  if (!located) return
  const { doc, tr } = view.state
  settleFormTextFields(doc, located.pos, located.pos + located.node.nodeSize, tr)
  if (tr.docChanged) view.dispatch(tr.setMeta(HOST_WRITE_META, true))
}
