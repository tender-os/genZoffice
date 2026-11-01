// Click-to-tick in review mode: a click on a box drawn as a text glyph toggles it, the way a reviewer
// expects a form to behave. It is an ordinary editor transaction, so it undoes and saves as typing does.
// Word's own checkbox controls and symbol-font boxes are not text glyphs and are left alone.
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'

const TOGGLED: Record<string, string> = { '☐': '☒', '☒': '☐', '☑': '☐' }

/** the box character at `pos` or just before it (a click lands on either side of a glyph), if any */
function boxAt(state: EditorState, pos: number) {
  const { doc } = state
  for (const from of [pos, pos - 1]) {
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
  const marks = state.doc.resolve(box.from + 1).marks()
  return state.tr.replaceWith(box.from, box.from + 1, state.schema.text(TOGGLED[box.char], marks))
}

export const boxTickKey = new PluginKey('voundBoxTick')

export const boxTickPlugin = new Plugin({
  key: boxTickKey,
  props: {
    handleClick(view, pos) {
      if (!view.editable) return false
      const toggle = toggleBoxAt(view.state, pos)
      if (!toggle) return false
      view.dispatch(toggle)
      return true
    },
  },
})
