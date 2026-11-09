// Click-to-tick in review mode: a click on a box drawn as a text glyph toggles it, the way a reviewer
// expects a form to behave; a click elsewhere on its line places the caret as usual. It is an ordinary editor transaction, so it undoes and saves as typing does.
// Word's legacy form boxes show as a glyph carrying the field's mark, which the save needs to keep the field.
// Word's other checkbox controls and symbol-font boxes are not text glyphs and are left alone.
import {
  Plugin,
  PluginKey,
  TextSelection,
  type EditorState,
  type Transaction,
} from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'

const TOGGLED: Record<string, string> = { '☐': '☒', '☒': '☐', '☑': '☐' }

/** the box character at `pos` or just before it (a click lands on either side of a glyph), if any */
function boxAt(state: EditorState, pos: number) {
  const { doc } = state
  // a click on a box that opens its line can report the position just before the line
  const inLine = doc.resolve(pos).nodeAfter?.isTextblock ? pos + 1 : pos
  for (const from of [inLine, inLine - 1]) {
    if (from < 0 || from + 1 > doc.content.size) continue
    const char = doc.textBetween(from, from + 1, '\u0000', '\u0000')
    if (char in TOGGLED) return { from, char }
  }
  return null
}

/** the transaction that toggles the box at `pos`; null when there is none. Its marks stay. */
export function toggleBoxAt(state: EditorState, pos: number): Transaction | null {
  const box = boxAt(state, pos)
  if (!box) return null
  const marks = state.doc.nodeAt(box.from)?.marks
  return state.tr.replaceWith(box.from, box.from + 1, state.schema.text(TOGGLED[box.char], marks))
}

/** how far from the glyph (px) a click still counts as on it */
const GLYPH_SLOP_PX = 2

/** the box the glyph character at `from` is drawn in; its element can be wider (boxGap.ts) */
function glyphRect(view: EditorView, from: number) {
  let { node, offset } = view.domAtPos(from, 1)
  while (node.nodeType !== Node.TEXT_NODE && node.childNodes[offset]) {
    node = node.childNodes[offset]
    offset = 0
  }
  const range = view.dom.ownerDocument.createRange()
  range.setStart(node, offset)
  range.setEnd(node, offset + 1)
  return range.getBoundingClientRect()
}

/** whether the click landed on the drawn glyph at `from`, not elsewhere on its line */
function clickedGlyph(view: EditorView, from: number, event: MouseEvent) {
  const glyph = glyphRect(view, from)
  return (
    event.clientX >= glyph.left - GLYPH_SLOP_PX &&
    event.clientX <= glyph.right + GLYPH_SLOP_PX &&
    event.clientY >= glyph.top - GLYPH_SLOP_PX &&
    event.clientY <= glyph.bottom + GLYPH_SLOP_PX
  )
}

export const boxTickKey = new PluginKey('voundBoxTick')

export const boxTickPlugin = new Plugin({
  key: boxTickKey,
  props: {
    handleClick(view, pos, event) {
      if (!view.editable) return false
      const box = boxAt(view.state, pos)
      const toggle = toggleBoxAt(view.state, pos)
      if (!box || !toggle || !clickedGlyph(view, box.from, event)) return false
      // the caret goes to the box's line, so typing does not continue wherever it was before the click
      view.dispatch(toggle.setSelection(TextSelection.create(toggle.doc, box.from + 1)))
      view.focus()
      return true
    },
  },
})
