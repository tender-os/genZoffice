// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import Document from '@tiptap/extension-document'
import Paragraph from '@tiptap/extension-paragraph'
import Text from '@tiptap/extension-text'
import { enterViewMode, leaveViewMode, viewModeKey } from '../viewMode'

let editor: Editor | null = null
let appListener: ((event: Event) => void) | null = null
afterEach(() => {
  if (editor) { leaveViewMode(editor); editor.view.dom.remove(); editor.destroy() }
  editor = null
  if (appListener) for (const type of ['keydown', 'dblclick', 'contextmenu']) document.removeEventListener(type, appListener, true)
  appListener = null
})

function viewedEditor() {
  const ed = new Editor({ element: document.createElement('div'), extensions: [Document, Paragraph, Text], content: '<p>Hallo Welt</p>' })
  editor = ed
  document.body.append(ed.view.dom)
  return ed
}

/** the app's earliest own listener: document capture, like App.tsx's keydown shield */
function appHears() {
  const heard = vi.fn()
  appListener = heard
  for (const type of ['keydown', 'dblclick', 'contextmenu']) document.addEventListener(type, heard, true)
  return heard
}

describe('view mode', () => {
  it('makes the editor read-only even when the app sets it editable', () => {
    const ed = viewedEditor()
    enterViewMode(ed)
    ed.setEditable(true)
    expect(ed.view.editable).toBe(false)
    expect(ed.view.dom.getAttribute('contenteditable')).toBe('false')
    expect(document.documentElement.classList.contains('vound-view')).toBe(true)
  })

  it('keeps keys, double-clicks and right-clicks from the app', () => {
    const ed = viewedEditor()
    const heard = appHears()
    enterViewMode(ed)
    ed.view.dom.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }))
    ed.view.dom.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }))
    ed.view.dom.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true }))
    expect(heard).not.toHaveBeenCalled()
  })

  it("stops Mod+S from opening the browser's save dialog but leaves copy alone", () => {
    const ed = viewedEditor()
    enterViewMode(ed)
    const save = new KeyboardEvent('keydown', { key: 's', metaKey: true, bubbles: true, cancelable: true })
    const copy = new KeyboardEvent('keydown', { key: 'c', metaKey: true, bubbles: true, cancelable: true })
    ed.view.dom.dispatchEvent(save)
    ed.view.dom.dispatchEvent(copy)
    expect(save.defaultPrevented).toBe(true)
    expect(copy.defaultPrevented).toBe(false)
  })

  it('leaves the editor as the app had it', () => {
    const ed = viewedEditor()
    const heard = appHears()
    enterViewMode(ed)
    enterViewMode(ed)
    leaveViewMode(ed)
    expect(viewModeKey.get(ed.state)).toBeUndefined()
    expect(ed.view.editable).toBe(true)
    expect(document.documentElement.classList.contains('vound-view')).toBe(false)
    ed.view.dom.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }))
    expect(heard).toHaveBeenCalledTimes(1)
  })
})
