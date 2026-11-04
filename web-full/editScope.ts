// Review mode can keep the reviewer's edits inside the fields, so a stray keystroke cannot change the form's
// own text. An edit outside the fields is refused whole and reported, and the host offers the way to unlock
// the document. Host writes, the app's own refreshes and a document opening are not the reviewer's edits, and
// a document with no field markers has nothing to protect the fields from, so it stays editable.
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import type { Decoration } from '@tiptap/pm/view'
import type { EditScope, EditorToHost } from './embed-protocol'
import { editedRanges, isUserEdit } from './fieldReporting'
import { markersKey } from './markersPlugin'

export const editScopeKey = new PluginKey('voundEditScope')

let scope: EditScope = 'everything'

/** what the reviewer may edit from now on (vound:editScope); it stays across saves and opens */
export const setEditScope = (next: EditScope) => {
  scope = next
  document.documentElement.classList.toggle('vound-fields-only', next === 'fields')
}

export const isLockedToFields = () => scope === 'fields'

const fieldBlocks = (state: EditorState) =>
  markersKey.getState(state)?.set.find().filter((decoration) => decoration.spec.isField) ?? []

/** an insertion is inside a block only strictly; a paragraph added between two blocks is not */
const isWithin = (block: Decoration, range: { from: number; to: number }) =>
  range.from === range.to
    ? block.from < range.from && range.from < block.to
    : block.from <= range.from && range.to <= block.to

const staysInFields = (tr: Transaction, blocks: Decoration[]) =>
  editedRanges([tr]).before.every((range) => blocks.some((block) => isWithin(block, range)))

/** the plugin refusing edits outside the fields, reported through `post`; one per review-mode session */
export const editScopePlugin = (post: (msg: EditorToHost) => void) =>
  new Plugin({
    key: editScopeKey,
    filterTransaction: (tr, state) => {
      if (scope === 'everything' || !isUserEdit(tr) || tr.getMeta('appendedTransaction')) return true
      const blocks = fieldBlocks(state)
      if (blocks.length === 0 || staysInFields(tr, blocks)) return true
      post({ type: 'vound:editRefused' })
      return false
    },
  })
