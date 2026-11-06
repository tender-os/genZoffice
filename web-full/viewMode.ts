// View mode (vound:open with mode 'view'): the host shows a document and offers nothing to change it.
// host.css hides the ribbon, its tab row and the status bar under html.vound-view and takes pointer input off
// the pages, so no click places a caret, ticks a box, drags a handle or opens a header. The editor takes no
// input, whatever the app's own setEditable says. Keys, double-clicks and right-clicks never reach the app's
// handlers (a window capture listener runs before all of them); scrolling stays.
import type { Editor } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'

export const viewModeKey = new PluginKey('voundViewMode')

/** ProseMirror treats the view as read-only once any plugin's editable answers false */
export const viewModePlugin = new Plugin({ key: viewModeKey, props: { editable: () => false } })

const KEPT_FROM_APP = ['keydown', 'dblclick', 'contextmenu'] as const

/** Mod+S, Mod+P and Mod+O would open the browser's save, print and open dialogs instead */
const BROWSER_DIALOG_KEYS = ['s', 'p', 'o']

function keepFromApp(event: Event) {
  event.stopPropagation()
  const opensBrowserDialog =
    event instanceof KeyboardEvent && (event.metaKey || event.ctrlKey) && BROWSER_DIALOG_KEYS.includes(event.key.toLowerCase())
  if (opensBrowserDialog) event.preventDefault()
}

/** Marks the page for host.css, keeps input from the app's handlers and makes the editor read-only. */
export function enterViewMode(editor: Editor): void {
  document.documentElement.classList.add('vound-view')
  for (const type of KEPT_FROM_APP) window.addEventListener(type, keepFromApp, true)
  if (!viewModeKey.get(editor.state)) editor.registerPlugin(viewModePlugin)
}

/** Back to the plain embedded editor, for a vound:open in another mode. */
export function leaveViewMode(editor: Editor): void {
  document.documentElement.classList.remove('vound-view')
  for (const type of KEPT_FROM_APP) window.removeEventListener(type, keepFromApp, true)
  if (viewModeKey.get(editor.state)) editor.unregisterPlugin(viewModeKey)
}
