// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import Bold from '@tiptap/extension-bold'
import Document from '@tiptap/extension-document'
import Italic from '@tiptap/extension-italic'
import Paragraph from '@tiptap/extension-paragraph'
import Text from '@tiptap/extension-text'
import Underline from '@tiptap/extension-underline'
import { baseKeymap } from '@tiptap/pm/commands'
import { keymap } from '@tiptap/pm/keymap'
import { EditorState, NodeSelection, TextSelection, type Plugin } from '@tiptap/pm/state'
import { EditorView, type EditorProps } from '@tiptap/pm/view'
import { mountFieldChip } from '../fieldChip'
import { fieldReportingKey } from '../fieldReporting'
import { enterReviewMode, FORM_FIELD_BLANK_CLASS, leaveReviewMode, reviewModeKey, reviewModePlugin } from '../reviewMode'
import { doc, fieldPara, para, schema } from './markerSchema'

// the chip itself is tested in fieldChip.test.ts; here only that review mode mounts and unmounts it
const unmountChip = vi.hoisted(() => vi.fn())
vi.mock('../fieldChip', () => ({ mountFieldChip: vi.fn(() => unmountChip) }))

// jsdom has no ClipboardEvent; ProseMirror's pasteHTML creates one when it is not given an event
beforeAll(() => {
  if (typeof globalThis.ClipboardEvent === 'undefined') {
    Object.defineProperty(globalThis, 'ClipboardEvent', { value: class extends Event { clipboardData = null }, configurable: true })
  }
})

let view: EditorView | null = null
let editor: Editor | null = null
afterEach(() => {
  view?.destroy(); view = null
  if (editor) { leaveReviewMode(editor); editor.view.dom.remove(); editor.destroy() }
  editor = null
  document.documentElement.classList.remove('vound-review')
})

const arial = schema.marks.docTextStyle.create({ font: 'Arial' })
/** one paragraph "Ort: " in Arial, caret at its end */
function reviewView(props: Partial<EditorProps> = {}) {
  const start = doc(schema.nodes.docParagraph.create(null, schema.text('Ort: ', [arial])))
  const state = EditorState.create({ schema, doc: start, plugins: [reviewModePlugin] })
  view = new EditorView(document.createElement('div'), {
    ...props,
    state: state.apply(state.tr.setSelection(TextSelection.create(state.doc, 6))),
  })
  return view
}
const runs = (v: EditorView) => {
  const out: { text: string; marks: unknown[] }[] = []
  v.state.doc.descendants((node) => { if (node.isText) out.push({ text: node.text ?? '', marks: node.marks.map((m) => m.attrs) }) })
  return out
}

describe('reviewModePlugin', () => {
  it("pasting HTML inserts its plain text with the insertion point's marks", () => {
    const v = reviewView()
    v.pasteHTML('<p><b>Berlin</b></p><p>Mitte</p>')
    expect(v.state.doc.childCount).toBe(1)
    expect(runs(v)).toEqual([{ text: 'Ort: Berlin\nMitte', marks: [{ bold: null, font: 'Arial' }] }])
  })

  it('a pasted line break survives as a line feed', () => {
    const v = reviewView()
    v.pasteHTML('<p>Berlin<br>Mitte</p>')
    expect(runs(v)).toEqual([{ text: 'Ort: Berlin\nMitte', marks: [{ bold: null, font: 'Arial' }] }])
  })

  it("a paste event is handled before the app's own paste handlers", () => {
    const appPaste = vi.fn(() => true)
    const v = reviewView({ handlePaste: appPaste })
    const event = new Event('paste', { bubbles: true, cancelable: true })
    const clipboard: Record<string, string> = { 'text/html': '<p><b>Berlin</b></p>', 'text/plain': 'Berlin\r\nMitte' }
    Object.defineProperty(event, 'clipboardData', { value: { getData: (type: string) => clipboard[type] ?? '' } })
    v.dom.dispatchEvent(event)
    expect(appPaste).not.toHaveBeenCalled()
    expect(event.defaultPrevented).toBe(true)
    expect(runs(v)).toEqual([{ text: 'Ort: Berlin\nMitte', marks: [{ bold: null, font: 'Arial' }] }])
  })

  it('an HTML-only clipboard pastes one line per block', () => {
    const v = reviewView()
    const event = new Event('paste', { bubbles: true, cancelable: true })
    const html = '<div>\n  <p><b>Berlin</b></p>\n  <ul><li>Mitte</li><li>Nord<br>Ost</li></ul>\n</div>'
    Object.defineProperty(event, 'clipboardData', { value: { getData: (type: string) => (type === 'text/html' ? html : '') } })
    v.dom.dispatchEvent(event)
    expect(runs(v)).toEqual([{ text: 'Ort: Berlin\nMitte\nNord\nOst', marks: [{ bold: null, font: 'Arial' }] }])
  })
})

// the platform's Mod key as prosemirror-keymap and tiptap resolve it
const mod = /Mac|iP(hone|[oa]d)/.test(navigator.platform) ? { metaKey: true } : { ctrlKey: true }
const press = (target: Editor, key: string, modifiers: KeyboardEventInit = mod) =>
  target.view.dom.dispatchEvent(new KeyboardEvent('keydown', { key, code: `Key${key.toUpperCase()}`, bubbles: true, cancelable: true, ...modifiers }))
const markNames = (target: Editor) => {
  const names = new Set<string>()
  target.state.doc.descendants((node) => { node.marks.forEach((m) => names.add(m.type.name)) })
  return [...names]
}

const post = vi.fn()

describe('enterReviewMode', () => {
  // a stand-in for App.tsx's window keydown handler: it edits on ⌘E without looking at defaultPrevented
  let appShortcuts: ((e: KeyboardEvent) => void) | null = null
  afterEach(() => { if (appShortcuts) window.removeEventListener('keydown', appShortcuts); appShortcuts = null })
  const tiptap = () => {
    const ed = new Editor({ element: document.createElement('div'), extensions: [Document, Paragraph, Text, Bold, Italic, Underline], content: '<p>Hallo Welt</p>' })
    editor = ed
    document.body.append(ed.view.dom)
    ed.commands.setTextSelection({ from: 1, to: 6 })
    appShortcuts = vi.fn((e: KeyboardEvent) => { if ((e.metaKey || e.ctrlKey) && e.key === 'e') ed.commands.insertContentAt(1, 'Mitte ') })
    window.addEventListener('keydown', appShortcuts)
    return ed
  }

  it('formatting shortcuts reach the editor keymaps and the app in review mode, document tools do not', () => {
    const ed = tiptap()
    enterReviewMode(ed, post)
    press(ed, 'b')
    expect(markNames(ed)).toEqual(['bold'])
    press(ed, 'e')
    expect(ed.getText()).toBe('Mitte Hallo Welt')
    expect(appShortcuts).toHaveBeenCalledTimes(2)
    // new comment ⌥⌘A and File › Open stay with the blocker
    press(ed, 'a', { ...mod, altKey: true })
    press(ed, 'o')
    expect(appShortcuts).toHaveBeenCalledTimes(2)
    leaveReviewMode(ed)
    press(ed, 'o')
    expect(appShortcuts).toHaveBeenCalledTimes(3)
  })

  it("the app's context menu opens in review mode", () => {
    const ed = tiptap()
    const appMenu = vi.fn()
    document.body.addEventListener('contextmenu', appMenu)
    enterReviewMode(ed, post)
    ed.view.dom.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }))
    expect(appMenu).toHaveBeenCalledTimes(1)
    document.body.removeEventListener('contextmenu', appMenu)
  })

  it('marks the page, registers the plugin once, and leaves again', () => {
    const ed = tiptap()
    enterReviewMode(ed, post)
    enterReviewMode(ed, post)
    expect(document.documentElement.classList.contains('vound-review')).toBe(true)
    expect(ed.state.plugins.filter((p) => p === reviewModePlugin)).toHaveLength(1)
    expect(ed.state.plugins.filter((p) => p.spec.key === fieldReportingKey)).toHaveLength(1)
    // Tab is the reviewer's "next field" in review mode: no indent, no tab character
    const tab = new KeyboardEvent('keydown', { key: 'Tab', code: 'Tab', bubbles: true, cancelable: true })
    ed.view.dom.dispatchEvent(tab)
    expect(tab.defaultPrevented).toBe(true)

    leaveReviewMode(ed)
    expect(document.documentElement.classList.contains('vound-review')).toBe(false)
    expect(reviewModeKey.get(ed.state)).toBeUndefined()
    expect(fieldReportingKey.get(ed.state)).toBeUndefined()
    press(ed, 'b')
    expect(markNames(ed)).toEqual(['bold'])
  })
})

describe('a field block selected as a whole in review mode', () => {
  /** "Firma:" and a form field selected as an object, as a click on it leaves it */
  function selectedField(plugins: Plugin[]) {
    const start = doc(para(0, 'Firma:'), fieldPara(1, 'Muster GmbH'))
    const state = EditorState.create({ schema, doc: start, plugins: [...plugins, keymap(baseKeymap)] })
    view = new EditorView(document.createElement('div'), { state: state.apply(state.tr.setSelection(NodeSelection.create(state.doc, start.child(0).nodeSize))) })
    return view
  }
  const fieldKept = (v: EditorView) => v.state.doc.childCount === 2 && v.state.doc.child(1).attrs.fieldDisplay?.left === 'Muster GmbH'
  const typeText = (v: EditorView, text: string) =>
    v.someProp('handleTextInput', (f) => f(v, v.state.selection.from, v.state.selection.to, text, () => v.state.tr.insertText(text)))
  const key = (v: EditorView, name: string) => {
    const event = new KeyboardEvent('keydown', { key: name, code: name, bubbles: true, cancelable: true })
    v.dom.dispatchEvent(event)
    return event
  }

  it('without review mode a key replaces the block (the control)', () => {
    const v = selectedField([])
    key(v, 'Backspace')
    expect(fieldKept(v)).toBe(false)
  })

  it('refuses typing, Backspace, Delete, paste and cut over it', () => {
    const v = selectedField([reviewModePlugin])
    expect(typeText(v, 'x')).toBe(true)
    expect(key(v, 'Backspace').defaultPrevented).toBe(true)
    expect(key(v, 'Delete').defaultPrevented).toBe(true)
    v.pasteHTML('<p>Beispiel</p>')
    for (const type of ['paste', 'cut']) {
      const event = new Event(type, { bubbles: true, cancelable: true })
      Object.defineProperty(event, 'clipboardData', { value: { getData: () => 'Beispiel', setData: () => {}, clearData: () => {} } })
      v.dom.dispatchEvent(event)
      expect(event.defaultPrevented).toBe(true)
    }
    expect(fieldKept(v)).toBe(true)
    expect(v.state.selection).toBeInstanceOf(NodeSelection)
  })

  it('leaves typing in a text selection alone', () => {
    const v = selectedField([reviewModePlugin])
    v.dispatch(v.state.tr.setSelection(TextSelection.create(v.state.doc, 7)))
    expect(typeText(v, 'x')).toBeFalsy()
    // a collapsed caret's Backspace is the browser's own deletion: nothing takes the key
    expect(key(v, 'Backspace').defaultPrevented).toBe(false)
  })
})

describe('the field chip in review mode', () => {
  it('is mounted once per editor on entering and unmounted on leaving', () => {
    const ed = new Editor({ element: document.createElement('div'), extensions: [Document, Paragraph, Text], content: '<p>Hallo</p>' })
    editor = ed
    vi.mocked(mountFieldChip).mockClear()
    unmountChip.mockClear()
    enterReviewMode(ed, post)
    enterReviewMode(ed, post)
    expect(mountFieldChip).toHaveBeenCalledTimes(1)
    expect(mountFieldChip).toHaveBeenCalledWith(ed.view, post)
    expect(unmountChip).not.toHaveBeenCalled()
    leaveReviewMode(ed)
    expect(unmountChip).toHaveBeenCalledTimes(1)
    // the afterEach leaves again: nothing left to unmount
    leaveReviewMode(ed)
    expect(unmountChip).toHaveBeenCalledTimes(1)
  })
})

describe("typing into Word's empty form field result in review mode", () => {
  const enSpaces = '\u2002'.repeat(5)
  function caretIn(text: string, offset: number) {
    const state = EditorState.create({ schema, doc: doc(para(0, 'Fax:'), para(1, text)), plugins: [reviewModePlugin] })
    const start = state.doc.child(0).nodeSize + 1
    view = new EditorView(document.createElement('div'), { state: state.apply(state.tr.setSelection(TextSelection.create(state.doc, start + offset))) })
    return view
  }
  const typeAtCaret = (v: EditorView, text: string) => {
    const { from, to } = v.state.selection
    const deflt = () => v.state.tr.insertText(text, from, to)
    if (!v.someProp('handleTextInput', (f) => f(v, from, to, text, deflt))) v.dispatch(deflt())
  }

  it('replaces the en spaces wherever the caret sits in them, then types on', () => {
    const v = caretIn(enSpaces, 5)
    typeAtCaret(v, '0')
    typeAtCaret(v, '8')
    expect(v.state.doc.child(1).textContent).toBe('08')
  })

  it('leaves ordinary spaces and text alone', () => {
    const spaces = caretIn('   ', 3)
    typeAtCaret(spaces, 'x')
    expect(spaces.state.doc.child(1).textContent).toBe('   x')
    const mixed = caretIn(`${enSpaces}Muster`, 5)
    typeAtCaret(mixed, 'x')
    expect(mixed.state.doc.child(1).textContent).toBe(`${enSpaces}xMuster`)
  })

  it('gives the blank no width, so the caret starts at the field edge, and keeps its en spaces', () => {
    const v = caretIn(enSpaces, 5)
    const blank = v.dom.querySelector(`.${FORM_FIELD_BLANK_CLASS}`)
    expect(blank?.textContent).toBe(enSpaces)
    expect(v.dom.querySelectorAll(`.${FORM_FIELD_BLANK_CLASS}`)).toHaveLength(1)
    typeAtCaret(v, '0')
    expect(v.dom.querySelector(`.${FORM_FIELD_BLANK_CLASS}`)).toBeNull()
  })
})
