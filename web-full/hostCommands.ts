// Undo, redo and zoom the host asks for (vound:command, vound:setZoom) run as the app's own menu commands, so they
// behave as the app's buttons do. The host hears what there is to undo or redo, and a frame the host gives focus
// back to hands it on to the editor, so Cmd+Z reaches the document again.
import { redoDepth, undoDepth } from '@tiptap/pm/history'
import { Plugin, PluginKey, type EditorState } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import type { EditorToHost, MenuHostCommand } from './embed-protocol'

/** the app's menu command (App.tsx, onMenuCommand) that each host command runs */
export const MENU_COMMANDS = {
  save: 'save',
  undo: 'undo',
  redo: 'redo',
  zoomIn: 'zoom-in',
  zoomOut: 'zoom-out',
  zoomFitWidth: 'zoom-page-width',
} as const satisfies Record<MenuHostCommand, string>

/** the app's menu command that sets the zoom to the percent given as its payload */
export const SET_ZOOM_COMMAND = 'zoom-set'

const historyOf = (state: EditorState) =>
  ({ type: 'vound:history', canUndo: undoDepth(state) > 0, canRedo: redoDepth(state) > 0 }) as const

export const historyReportKey = new PluginKey('voundHistoryReport')

/** tells the host whether there is something to undo or redo: once when registered, then on each change */
export const historyReportPlugin = (post: (msg: EditorToHost) => void) =>
  new Plugin({
    key: historyReportKey,
    view(view) {
      let last = historyOf(view.state)
      post(last)
      return {
        update(updated) {
          const next = historyOf(updated.state)
          if (next.canUndo === last.canUndo && next.canRedo === last.canRedo) return
          last = next
          post(next)
        },
      }
    },
  })

/** focuses the editor when the frame gets focus back with nothing in it focused; returns the unsubscribe */
export function refocusOnReturn(viewOf: () => EditorView | null) {
  const onFocus = () => {
    const view = viewOf()
    const nothingFocused =
      document.activeElement === null || document.activeElement === document.body
    if (view && !view.isDestroyed && nothingFocused) view.focus()
  }
  window.addEventListener('focus', onFocus)
  return () => window.removeEventListener('focus', onFocus)
}
