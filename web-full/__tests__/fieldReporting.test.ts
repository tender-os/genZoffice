// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { EditorState, NodeSelection, Plugin, TextSelection } from '@tiptap/pm/state'
import { EditorView } from '@tiptap/pm/view'
import type { EditorToHost, SlotMarker } from '../embed-protocol'
import { fieldReportingPlugin } from '../fieldReporting'
import { applyFieldValue, HOST_WRITE_META, placeCaret } from '../fieldWrite'
import { markersPlugin, setMarkers } from '../markersPlugin'
import { fieldNodeView } from './fieldNodeView'
import { doc, fieldPara, para, paragraphs, schema } from './markerSchema'

const marker = (id: string, paragraph: number): SlotMarker => ({ id, paragraph, scope: 'paragraph', tone: 'gap', label: id, actions: [] })

let view: EditorView | null = null
afterEach(() => { view?.dom.remove(); view?.destroy(); view = null })

/** paragraphs 0..4; "ort" and "datum" share paragraph 1, "name" marks paragraph 3 */
function reportingView() {
  const post = vi.fn<(msg: EditorToHost) => void>()
  const state = EditorState.create({ schema, doc: doc(...paragraphs(0, 5)), plugins: [markersPlugin, fieldReportingPlugin(post)] })
  view = new EditorView(document.createElement('div'), { state })
  setMarkers(view, [marker('ort', 1), marker('datum', 1), marker('name', 3)])
  return { v: view, post }
}
/** the position inside the paragraph with this source index, `offset` characters in */
const inside = (v: EditorView, index: number, offset = 0) => {
  let found = -1
  v.state.doc.forEach((node, pos) => { if (node.attrs.docxPara === index) found = pos + 1 + offset })
  return found
}
const caretTo = (v: EditorView, pos: number) => v.dispatch(v.state.tr.setSelection(TextSelection.create(v.state.doc, pos)))
const pressTab = (v: EditorView, init: KeyboardEventInit = {}, target: Element = v.dom) => {
  const event = new KeyboardEvent('keydown', { key: 'Tab', code: 'Tab', bubbles: true, cancelable: true, ...init })
  target.dispatchEvent(event)
  return event
}
const sent = (post: ReturnType<typeof vi.fn>, type: EditorToHost['type']) =>
  post.mock.calls.map(([msg]) => msg).filter((msg: EditorToHost) => msg.type === type)

describe('fieldReportingPlugin', () => {
  it('reports blockEdited for a user edit inside a marked paragraph', () => {
    const { v, post } = reportingView()
    v.dispatch(v.state.tr.insertText('Berlin', inside(v, 1, 3)))
    expect(sent(post, 'vound:blockEdited')).toEqual([{ type: 'vound:blockEdited', ids: ['ort', 'datum'] }])

    // one transaction over two marked paragraphs reports each id once
    post.mockClear()
    v.dispatch(v.state.tr.insertText('x', inside(v, 3, 1)).insertText('y', inside(v, 1, 1)).insertText('z', inside(v, 1, 2)))
    expect(sent(post, 'vound:blockEdited')).toEqual([{ type: 'vound:blockEdited', ids: ['ort', 'datum', 'name'] }])
  })

  it('does not report a host write', () => {
    const { v, post } = reportingView()
    expect(applyFieldValue(v, { block: { paragraph: 1, scope: 'paragraph' }, previous: 'paragraph 1', value: 'Berlin' }).ok).toBe(true)
    expect(v.state.doc.child(1).textContent).toBe('Berlin')
    expect(sent(post, 'vound:blockEdited')).toEqual([])
  })

  it('does not report an edit outside markers', () => {
    const { v, post } = reportingView()
    v.dispatch(v.state.tr.insertText('x', inside(v, 2, 1)))
    // a paragraph inserted between two blocks touches the marked neighbours only at their edges
    v.dispatch(v.state.tr.insert(inside(v, 4) - 1, para(null, 'new')))
    // so does deleting the unmarked paragraph between the marked paragraphs 1 and 3
    const second = inside(v, 2) - 1
    v.dispatch(v.state.tr.delete(second, second + v.state.doc.nodeAt(second)!.nodeSize))
    // a change that does not keep history (a field refresh, a pagination stamp) is the app's, not the reviewer's
    v.dispatch(v.state.tr.insertText('q', inside(v, 1, 1)).setMeta('addToHistory', false))
    // a whole-document replace opens another document
    v.dispatch(v.state.tr.replaceWith(0, v.state.doc.content.size, paragraphs(0, 5)))
    expect(sent(post, 'vound:blockEdited')).toEqual([])
  })

  it('reports caretBlock on entering and leftWith Tab on leaving by Tab', () => {
    const { v, post } = reportingView()
    caretTo(v, inside(v, 0, 2))
    expect(sent(post, 'vound:caretBlock')).toEqual([])

    caretTo(v, inside(v, 1, 2))
    // moving inside the block is not news
    caretTo(v, inside(v, 1, 4))
    expect(sent(post, 'vound:caretBlock')).toEqual([{ type: 'vound:caretBlock', ids: ['ort', 'datum'] }])

    const before = v.state.doc
    const tab = pressTab(v)
    // Tab neither indents nor inserts a tab in review mode: it is the reviewer's "next field"
    expect(tab.defaultPrevented).toBe(true)
    expect(v.state.doc.eq(before)).toBe(true)
    expect(sent(post, 'vound:caretBlock').at(-1)).toEqual({ type: 'vound:caretBlock', ids: [], leftWith: 'Tab' })

    // a later move into another marked block enters it; leaving it by a click is 'other'
    caretTo(v, inside(v, 3, 1))
    caretTo(v, inside(v, 4, 1))
    expect(sent(post, 'vound:caretBlock').slice(-2)).toEqual([
      { type: 'vound:caretBlock', ids: ['name'] },
      { type: 'vound:caretBlock', ids: [], leftWith: 'other' },
    ])

    // Tab outside a marked block is swallowed and reports nothing
    post.mockClear()
    expect(pressTab(v).defaultPrevented).toBe(true)
    expect(post).not.toHaveBeenCalled()
  })

  it('reports the block the caret is in when the markers arrive', () => {
    const post = vi.fn<(msg: EditorToHost) => void>()
    const state = EditorState.create({ schema, doc: doc(...paragraphs(0, 3)), plugins: [markersPlugin, fieldReportingPlugin(post)] })
    view = new EditorView(document.createElement('div'), { state })
    caretTo(view, inside(view, 2, 1))
    setMarkers(view, [marker('name', 2)])
    expect(sent(post, 'vound:caretBlock')).toEqual([{ type: 'vound:caretBlock', ids: ['name'] }])
  })

  it('Shift+Tab does nothing and reports nothing', () => {
    const { v, post } = reportingView()
    caretTo(v, inside(v, 1, 2))
    post.mockClear()
    const before = v.state.doc
    expect(pressTab(v, { shiftKey: true }).defaultPrevented).toBe(true)
    expect(v.state.doc.eq(before)).toBe(true)
    expect(post).not.toHaveBeenCalled()
  })

  it('reports a marked paragraph joined into its predecessor', () => {
    const { v, post } = reportingView()
    // Backspace at the start of paragraph 3 ("name") joins it into paragraph 2; its marker goes with it
    v.dispatch(v.state.tr.join(inside(v, 3) - 1))
    expect(v.state.doc.childCount).toBe(4)
    expect(sent(post, 'vound:blockEdited')).toEqual([{ type: 'vound:blockEdited', ids: ['name'] }])
  })

  it('does not report a transaction appended to a host write', () => {
    const post = vi.fn<(msg: EditorToHost) => void>()
    // a plugin that stamps something into the written paragraph after every host write, as app plugins may
    const stamp = new Plugin({
      appendTransaction: (trs, _old, state) => (trs.some((tr) => tr.getMeta(HOST_WRITE_META)) ? state.tr.insertText('!', 2) : null),
    })
    const state = EditorState.create({ schema, doc: doc(...paragraphs(0, 3)), plugins: [markersPlugin, stamp, fieldReportingPlugin(post)] })
    view = new EditorView(document.createElement('div'), { state })
    setMarkers(view, [marker('name', 0)])
    expect(applyFieldValue(view, { block: { paragraph: 0, scope: 'paragraph' }, previous: 'paragraph 0', value: 'Muster' }).ok).toBe(true)
    expect(view.state.doc.child(0).textContent).toBe('M!uster')
    expect(sent(post, 'vound:blockEdited')).toEqual([])
  })

  it('forgets the reported block when another document replaces the whole document', () => {
    const { v, post } = reportingView()
    caretTo(v, inside(v, 1, 2))
    // the opened document arrives (its markers are rebuilt on it, markersPlugin.ts)
    v.dispatch(v.state.tr.replaceWith(0, v.state.doc.content.size, paragraphs(0, 5)))
    post.mockClear()
    // the caret in the same paragraph of the new document is news to the host
    caretTo(v, inside(v, 1, 2))
    expect(sent(post, 'vound:caretBlock')).toEqual([{ type: 'vound:caretBlock', ids: ['ort', 'datum'] }])
  })
})

describe('fieldReportingPlugin in a field passthrough', () => {
  /** a field passthrough (paragraph 1, "firma") between two paragraphs, rendered by the field node view */
  function fieldView() {
    const post = vi.fn<(msg: EditorToHost) => void>()
    const mount = document.createElement('div')
    document.body.append(mount)
    const state = EditorState.create({
      schema,
      doc: doc(para(0, 'Firma:'), fieldPara(1, 'Muster GmbH'), para(2, 'Ort:')),
      plugins: [markersPlugin, fieldReportingPlugin(post)],
    })
    view = new EditorView(mount, { state, nodeViews: { docProtected: fieldNodeView } })
    setMarkers(view, [marker('firma', 1)])
    return { v: view, post }
  }

  it('reports the caret in a field being edited, commits on Tab, then reports the edit and leaving by Tab', () => {
    const { v, post } = fieldView()
    expect(placeCaret(v, { paragraph: 1, scope: 'paragraph' })).toBe(true)
    expect(sent(post, 'vound:caretBlock')).toEqual([{ type: 'vound:caretBlock', ids: ['firma'] }])

    // typing inside the node view: the DOM changes, the document only on commit
    const text = v.dom.querySelector('.doc-field-text')!
    text.textContent = 'Beispiel AG'
    post.mockClear()
    const tab = pressTab(v, {}, text)
    // the node view keeps ProseMirror's own keydown handling out; the browser would move the focus
    expect(tab.defaultPrevented).toBe(true)
    expect(v.state.doc.child(1).attrs.fieldDisplay.left).toBe('Beispiel AG')
    expect(post.mock.calls.map(([msg]) => msg)).toEqual([
      { type: 'vound:blockEdited', ids: ['firma'] },
      { type: 'vound:caretBlock', ids: [], leftWith: 'Tab' },
    ])
  })

  it('Shift+Tab in a field being edited keeps the focus and reports nothing', () => {
    const { v, post } = fieldView()
    placeCaret(v, { paragraph: 1, scope: 'paragraph' })
    post.mockClear()
    expect(pressTab(v, { shiftKey: true }, v.dom.querySelector('.doc-field-text')!).defaultPrevented).toBe(true)
    expect(post).not.toHaveBeenCalled()
  })

  it('a field selected as an object counts as the caret in its block', () => {
    const { v, post } = fieldView()
    v.dispatch(v.state.tr.setSelection(NodeSelection.create(v.state.doc, v.state.doc.child(0).nodeSize)))
    expect(sent(post, 'vound:caretBlock')).toEqual([{ type: 'vound:caretBlock', ids: ['firma'] }])
  })
})
