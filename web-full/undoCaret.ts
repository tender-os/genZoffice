// An undo or redo restores the selection from before the step it takes back, which is where the reviewer stood
// then, not the field the step changed: after "next field, Reject", Undo would put the caret in the previous field
// and the host would follow it there. In review mode the caret goes into the marked block the step changed. Only a
// text caret in a paragraph, never into a block shown whole (a field passthrough), whose own view keeps its text.
import { combineTransactionSteps, getChangedRanges } from '@tiptap/core'
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state'
import { fieldBlocks } from './editScope'

export const undoCaretKey = new PluginKey('voundUndoCaret')

export const undoCaretPlugin = new Plugin({
  key: undoCaretKey,
  appendTransaction(transactions, oldState, state) {
    if (!transactions.some((tr) => tr.getMeta('history$'))) return null
    const blocks = fieldBlocks(state)
    const { selection } = state
    for (const { newRange } of getChangedRanges(
      combineTransactionSteps(oldState.doc, [...transactions]),
    )) {
      const block = blocks.find((field) => field.from <= newRange.from && newRange.to <= field.to)
      if (!block) continue
      if (block.from <= selection.from && selection.to <= block.to) return null
      if (!state.doc.resolve(newRange.to).parent.isTextblock) continue
      return state.tr.setSelection(TextSelection.create(state.doc, newRange.to))
    }
    return null
  },
})
