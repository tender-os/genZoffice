// Review mode (vound:open with mode 'review'): the host shows its own bar instead of the ribbon, so the
// editor works like a text field. Paste inserts plain text in the insertion point's formatting; the
// formatting and document-tool shortcuts do nothing (reviewShortcuts.ts); the app's own context menu does
// not open, so its Paste entries cannot bypass the plain-text rule (the browser's menu opens instead, and
// its Paste is an ordinary paste event). host.css hides the ribbon, its tab row and the status bar under
// html.vound-review. Edits and caret moves in marked blocks are reported to the host (fieldReporting.ts), and
// hovering a marked block shows its field chip (fieldChip.ts).
// A form field selected as a whole block (a click on it selects it as an object) cannot be typed, deleted,
// pasted or cut over: that would replace the block, and with it the field, by plain text. Typing into
// Word's empty form field result (en spaces) replaces the blank instead of adding to it.
import type { Editor } from '@tiptap/core'
import { NodeSelection, Plugin, PluginKey, type EditorState } from '@tiptap/pm/state'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { EditorView } from '@tiptap/pm/view'
import type { EditorToHost } from './embed-protocol'
import { boxTickKey, boxTickPlugin } from './boxTick'
import { mountFieldChip } from './fieldChip'
import { fieldReportingKey, fieldReportingPlugin } from './fieldReporting'
import { clipboardText } from './pasteText'
import { blockedInReview, isMacPlatform, keepsNativeAction } from './reviewShortcuts'

export const reviewModeKey = new PluginKey('voundReviewMode')

/** Plain text at the selection; insertText keeps the stored marks or those at the insertion point. */
function insertPlainText(view: EditorView, text: string) {
  if (text) view.dispatch(view.state.tr.insertText(text).scrollIntoView())
  return true
}

/** a field passthrough (fieldDisplay) selected as a whole block */
export const selectsFieldBlock = (state: EditorState) =>
  state.selection instanceof NodeSelection && state.selection.node.attrs.fieldDisplay != null

/** swallow an event that would replace a selected field block */
function keepFieldBlock(view: EditorView, event: Event) {
  if (!selectsFieldBlock(view.state)) return false
  event.preventDefault()
  return true
}

const deletesSelection = (event: KeyboardEvent) => event.key === 'Backspace' || event.key === 'Delete'

/** A paragraph of plain text made of en spaces only. That is how Word's empty legacy form field result shows
 *  in a table cell (a plain paragraph holding the field's result), but the editor does not mark field
 *  paragraphs, so any such paragraph matches: en spaces are not typed by hand. */
function isFormFieldBlank(node: PmNode) {
  let plainText = node.childCount > 0
  node.forEach((child) => { plainText &&= child.isText })
  return plainText && /^\u2002+$/.test(node.textContent)
}

/** Typing into an empty form field result replaces its en spaces, wherever the caret sits in them. */
function replaceFormFieldBlank(view: EditorView, from: number, to: number, text: string) {
  const $from = view.state.doc.resolve(from)
  if (!$from.sameParent(view.state.doc.resolve(to)) || !isFormFieldBlank($from.parent)) return false
  view.dispatch(view.state.tr.insertText(text, $from.start(), $from.end()))
  return true
}

export const reviewModePlugin = new Plugin({
  key: reviewModeKey,
  props: {
    // The app's own paste handlers are direct editor props, which ProseMirror asks before any plugin,
    // and some of them insert formatted content themselves. The DOM event reaches this plugin first.
    handleDOMEvents: {
      paste: (view, event) => {
        if (keepFieldBlock(view, event)) return true
        if (!view.editable || !event.clipboardData) return false
        event.preventDefault()
        return insertPlainText(view, clipboardText(event.clipboardData))
      },
      cut: keepFieldBlock,
    },
    // A paste that skips the DOM event (view.pasteHTML) meets the app's direct handlePaste first and lands
    // here only when that one declines. Review mode closes the app's own ways in (the paste-options chip is
    // hidden, its context menu does not open); this keeps any remaining one plain. Hard breaks become line feeds.
    handlePaste: (view, _event, slice) =>
      selectsFieldBlock(view.state) || insertPlainText(view, slice.content.textBetween(0, slice.content.size, '\n', '\n')),
    handleTextInput: (view, from, to, text) => selectsFieldBlock(view.state) || replaceFormFieldBlank(view, from, to, text),
    handleKeyPress: (view) => selectsFieldBlock(view.state),
    handleKeyDown: (view, event) => deletesSelection(event) && selectsFieldBlock(view.state),
  },
})

// Window capture runs before every other listener: the editors' keymaps (the main editor, header and footer,
// text boxes) and App.tsx's window keydown handler, which acts without looking at defaultPrevented.
function blockShortcut(event: KeyboardEvent) {
  const mac = isMacPlatform()
  if (!blockedInReview(event, mac)) return
  event.stopPropagation()
  if (!keepsNativeAction(event, mac)) event.preventDefault()
}
// The app opens its menu from a React contextmenu handler; stopping the event before it keeps the browser's.
const keepAppMenuShut = (event: MouseEvent) => event.stopPropagation()

/** the field chip of the editor in review mode, mounted once per editor view */
let chip: { view: EditorView; unmount: () => void } | null = null

/** ahead of every other plugin: their keymaps and paste handlers would act first otherwise */
const first = (plugin: Plugin, plugins: Plugin[]) => [plugin, ...plugins]

/** Marks the page for host.css, registers the plugins once, blocks the shortcuts and the app's menu,
 *  reports the reviewer's edits and caret moves in marked blocks through `post`, and mounts the field chip,
 *  whose presses go through `post` too. */
export function enterReviewMode(editor: Editor, post: (msg: EditorToHost) => void): void {
  document.documentElement.classList.add('vound-review')
  // the same listener added twice is one listener
  window.addEventListener('keydown', blockShortcut, true)
  window.addEventListener('contextmenu', keepAppMenuShut, true)
  if (!reviewModeKey.get(editor.state)) editor.registerPlugin(reviewModePlugin, first)
  if (!fieldReportingKey.get(editor.state)) editor.registerPlugin(fieldReportingPlugin(post), first)
  if (!boxTickKey.get(editor.state)) editor.registerPlugin(boxTickPlugin, first)
  if (chip?.view === editor.view) return
  chip?.unmount()
  chip = { view: editor.view, unmount: mountFieldChip(editor.view, post) }
}

/** Back to the plain embedded editor, for a vound:open without review mode. */
export function leaveReviewMode(editor: Editor): void {
  document.documentElement.classList.remove('vound-review')
  window.removeEventListener('keydown', blockShortcut, true)
  window.removeEventListener('contextmenu', keepAppMenuShut, true)
  if (reviewModeKey.get(editor.state)) editor.unregisterPlugin(reviewModeKey)
  if (fieldReportingKey.get(editor.state)) editor.unregisterPlugin(fieldReportingKey)
  if (boxTickKey.get(editor.state)) editor.unregisterPlugin(boxTickKey)
  chip?.unmount()
  chip = null
}
