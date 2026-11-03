// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { EditorState } from '@tiptap/pm/state'
import { EditorView } from '@tiptap/pm/view'
import { editScopePlugin, setEditScope } from '../editScope'
import type { EditorToHost, SlotMarker } from '../embed-protocol'
import { HOST_WRITE_META } from '../fieldWrite'
import { markersPlugin, setMarkers } from '../markersPlugin'
import { doc, paragraphs, schema } from './markerSchema'

const field = (paragraph: number): SlotMarker => ({ id: `f${paragraph}`, paragraph, scope: 'paragraph', tone: 'gap', label: 'f', actions: ['fill'] })
const changeMarker = (paragraph: number): SlotMarker => ({ ...field(paragraph), id: `c${paragraph}`, actions: [] })

let view: EditorView | null = null
afterEach(() => {
  view?.destroy()
  view = null
  setEditScope('everything')
})

/** paragraphs 0..3 with the given markers, locked to the fields */
function lockedView(markers: SlotMarker[]) {
  const post = vi.fn<(msg: EditorToHost) => void>()
  const state = EditorState.create({ schema, doc: doc(...paragraphs(0, 4)), plugins: [markersPlugin, editScopePlugin(post)] })
  view = new EditorView(document.createElement('div'), { state })
  setMarkers(view, markers)
  setEditScope('fields')
  return { v: view, post }
}
const start = (v: EditorView, index: number) => {
  let found = -1
  v.state.doc.forEach((node, pos) => { if (node.attrs.docxPara === index) found = pos })
  return found
}
const textOf = (v: EditorView, index: number) => v.state.doc.child(index).textContent
const typeAt = (v: EditorView, index: number) => v.dispatch(v.state.tr.insertText('x', start(v, index) + 2))

describe('editScopePlugin', () => {
  it('takes an edit inside a field', () => {
    const { v, post } = lockedView([field(1)])
    typeAt(v, 1)
    expect(textOf(v, 1)).toContain('x')
    expect(post).not.toHaveBeenCalled()
  })

  it('refuses an edit outside the fields and reports it', () => {
    const { v, post } = lockedView([field(1)])
    const before = v.state.doc
    typeAt(v, 0)
    expect(v.state.doc.eq(before)).toBe(true)
    expect(post).toHaveBeenCalledWith({ type: 'vound:editRefused' })
  })

  it('refuses a deletion that starts in a field and runs past it', () => {
    const { v } = lockedView([field(1)])
    const before = v.state.doc
    v.dispatch(v.state.tr.delete(start(v, 1) + 2, start(v, 2) + 2))
    expect(v.state.doc.eq(before)).toBe(true)
  })

  it('refuses a paragraph added right before a field', () => {
    const { v } = lockedView([field(1)])
    const before = v.state.doc
    v.dispatch(v.state.tr.insert(start(v, 1), schema.nodes.docParagraph.create()))
    expect(v.state.doc.eq(before)).toBe(true)
  })

  it('does not treat a marked change outside the fields as a field', () => {
    const { v } = lockedView([field(1), changeMarker(2)])
    const before = v.state.doc
    typeAt(v, 2)
    expect(v.state.doc.eq(before)).toBe(true)
  })

  it('takes the host writes and the app refreshes anywhere', () => {
    const { v } = lockedView([field(1)])
    v.dispatch(v.state.tr.insertText('h', start(v, 0) + 2).setMeta(HOST_WRITE_META, true))
    expect(textOf(v, 0)).toContain('h')
    v.dispatch(v.state.tr.insertText('r', start(v, 3) + 2).setMeta('addToHistory', false))
    expect(textOf(v, 3)).toContain('r')
  })

  it('takes every edit with the scope everything', () => {
    const { v, post } = lockedView([field(1)])
    setEditScope('everything')
    typeAt(v, 0)
    expect(textOf(v, 0)).toContain('x')
    expect(post).not.toHaveBeenCalled()
  })

  it('keeps a document with no field markers editable', () => {
    const { v } = lockedView([changeMarker(2)])
    typeAt(v, 0)
    expect(textOf(v, 0)).toContain('x')
  })
})
