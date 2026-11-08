// Placeholder text (the editor's ctrlPlaceholder mark: a content control showing its placeholder, a form field's
// default text) is selected whole when the caret lands in it, as in Word, so the first keystroke replaces it and
// typing never mixes with it. A click anywhere on it selects it, its edges too, before the click ends, so a key
// pressed right after it cannot reach the browser's own caret; a caret moved by keys only from inside, so the
// reviewer can still step past it. Text without the mark is never selected, whatever it looks like.
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'

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

/** a click on placeholder text selects all of it; true when it did */
function selectClicked(view: EditorView, pos: number) {
  const range = placeholderAt(view.state.doc, pos)
  if (!range) return false
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, range.from, range.to)))
  return true
}

export const placeholderSelectKey = new PluginKey('voundPlaceholderSelect')

export const placeholderSelectPlugin = new Plugin({
  key: placeholderSelectKey,
  appendTransaction: (transactions, _oldState, state) => {
    const { selection } = state
    if (!selection.empty || !transactions.some((tr) => tr.selectionSet)) return null
    const range = placeholderAt(state.doc, selection.from)
    if (!range || selection.from === range.from || selection.from === range.to) return null
    return state.tr.setSelection(TextSelection.create(state.doc, range.from, range.to))
  },
  props: { handleClick: selectClicked },
})
