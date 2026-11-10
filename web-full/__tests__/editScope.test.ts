// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { EditorState, NodeSelection, TextSelection } from '@tiptap/pm/state'
import { Slice } from '@tiptap/pm/model'
import { EditorView } from '@tiptap/pm/view'
import { editScopePlugin, setEditScope } from '../editScope'
import type { EditorToHost, SlotMarker } from '../embed-protocol'
import { HOST_WRITE_META } from '../fieldWrite'
import { markersPlugin, setMarkers } from '../markersPlugin'
import { doc, fieldPara, para, paragraphs, protectedBlock, schema } from './markerSchema'

const field = (paragraph: number): SlotMarker => ({
  id: `f${paragraph}`,
  paragraph,
  scope: 'paragraph',
  tone: 'gap',
  label: 'f',
  actions: ['fill'],
})
const changeMarker = (paragraph: number): SlotMarker => ({
  ...field(paragraph),
  id: `c${paragraph}`,
  actions: [],
})

let view: EditorView | null = null
afterEach(() => {
  view?.destroy()
  view = null
  setEditScope('everything')
})

/** paragraphs 0..3 with the given markers, locked to the fields */
function lockedView(markers: SlotMarker[]) {
  const post = vi.fn<(msg: EditorToHost) => void>()
  const state = EditorState.create({
    schema,
    doc: doc(...paragraphs(0, 4)),
    plugins: [markersPlugin, editScopePlugin(post)],
  })
  view = new EditorView(document.createElement('div'), { state })
  setMarkers(view, markers)
  setEditScope('fields')
  return { v: view, post }
}
const start = (v: EditorView, index: number) => {
  let found = -1
  v.state.doc.forEach((node, pos) => {
    if (node.attrs.docxPara === index) found = pos
  })
  return found
}
const textOf = (v: EditorView, index: number) => v.state.doc.child(index).textContent
const typeAt = (v: EditorView, index: number) =>
  v.dispatch(v.state.tr.insertText('x', start(v, index) + 2))

describe('editScopePlugin', () => {
  it('keeps the words on a box line, and lets the box tick', () => {
    const { v, post } = lockedView([{ ...field(1), entry: 'toggle' }])
    v.dispatch(v.state.tr.insertText('☐', start(v, 1) + 1).setMeta(HOST_WRITE_META, true))
    const before = v.state.doc
    typeAt(v, 1)
    expect(v.state.doc.eq(before)).toBe(true)
    expect(post).toHaveBeenCalledWith({ type: 'vound:editRefused', where: 'boxLine' })

    v.dispatch(v.state.tr.insertText('☒', start(v, 1) + 1, start(v, 1) + 2))
    expect(textOf(v, 1)).toContain('☒')
  })

  it('leaves a line with a box and a text field open to typing', () => {
    const { v } = lockedView([
      { ...field(1), entry: 'toggle' },
      { ...field(1), id: 'text' },
    ])
    typeAt(v, 1)
    expect(textOf(v, 1)).toContain('x')
  })

  it('takes an edit inside a field', () => {
    const { v, post } = lockedView([field(1)])
    typeAt(v, 1)
    expect(textOf(v, 1)).toContain('x')
    expect(post).not.toHaveBeenCalled()
  })

  it("takes edits in the reviewer's own text, which offers no action", () => {
    const { v, post } = lockedView([field(2), { ...changeMarker(1), entry: 'ownText' }])
    typeAt(v, 1)
    expect(textOf(v, 1)).toContain('x')
    expect(post).not.toHaveBeenCalled()
  })

  it('refuses an edit outside the fields and reports it', () => {
    const { v, post } = lockedView([field(1)])
    const before = v.state.doc
    typeAt(v, 0)
    expect(v.state.doc.eq(before)).toBe(true)
    expect(post).toHaveBeenCalledWith({ type: 'vound:editRefused', where: 'outsideFields' })
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

  it('marks the page only while locked to the fields', () => {
    const locked = () => document.documentElement.classList.contains('vound-fields-only')
    setEditScope('fields')
    expect(locked()).toBe(true)
    setEditScope('everything')
    expect(locked()).toBe(false)
  })

  it('classes a field block, not a marked change', () => {
    const { v } = lockedView([field(1), changeMarker(2)])
    expect(v.dom.querySelectorAll('.vound-slot-field')).toHaveLength(1)
    expect(v.dom.querySelectorAll('.vound-slot')).toHaveLength(2)
  })

  it('keeps a document with no field markers editable', () => {
    const { v } = lockedView([changeMarker(2)])
    typeAt(v, 0)
    expect(textOf(v, 0)).toContain('x')
  })

  describe('a block shown as a chip, which takes no text', () => {
    /** a label line, an opaque field chip, the next line: the chip carries the field marker */
    function chipView(scope: 'fields' | 'everything') {
      const post = vi.fn<(msg: EditorToHost) => void>()
      const blocks = [
        para(0, 'Frau/Herr'),
        protectedBlock(1),
        para(2, 'ist heute verpflichtet worden.'),
      ]
      const state = EditorState.create({
        schema,
        doc: doc(...blocks),
        plugins: [markersPlugin, editScopePlugin(post)],
      })
      view = new EditorView(document.createElement('div'), { state })
      setMarkers(view, [field(1)])
      setEditScope(scope)
      return { v: view, post }
    }
    const selectChip = (v: EditorView) =>
      v.dispatch(v.state.tr.setSelection(NodeSelection.create(v.state.doc, start(v, 1))))

    for (const scope of ['fields', 'everything'] as const) {
      it(`refuses typing, pasting and deleting on the selected chip (${scope})`, () => {
        const { v, post } = chipView(scope)
        const before = v.state.doc
        selectChip(v)
        v.dispatch(v.state.tr.insertText('are be Limonka'))
        v.dispatch(v.state.tr.replaceSelection(new Slice(doc(para(null, 'Limonka')).content, 0, 0)))
        v.dispatch(v.state.tr.deleteSelection())
        expect(v.state.doc.eq(before)).toBe(true)
        expect(post).toHaveBeenCalledWith({ type: 'vound:editRefused', where: 'outsideFields' })
      })
    }

    it('refuses deleting a selection that runs from the text over the chip', () => {
      const { v } = chipView('everything')
      const before = v.state.doc
      v.dispatch(
        v.state.tr.setSelection(
          TextSelection.create(v.state.doc, start(v, 0) + 3, start(v, 2) + 4),
        ),
      )
      v.dispatch(v.state.tr.deleteSelection())
      expect(v.state.doc.eq(before)).toBe(true)
    })

    it('takes typing next to the chip and a paragraph added beside it', () => {
      const { v, post } = chipView('everything')
      typeAt(v, 2)
      v.dispatch(v.state.tr.insert(start(v, 1), schema.nodes.docParagraph.create()))
      expect(v.state.doc.textContent).toContain('x')
      expect(v.state.doc.childCount).toBe(4)
      expect(post).not.toHaveBeenCalled()
    })

    it("takes an in-place edit of the chip's own attributes (a nested table's cells, a text box)", () => {
      const { v, post } = chipView('everything')
      const pos = start(v, 1)
      v.dispatch(
        v.state.tr.setNodeMarkup(pos, undefined, {
          ...v.state.doc.nodeAt(pos)!.attrs,
          styleId: 'Bearbeitet',
        }),
      )
      expect(v.state.doc.nodeAt(pos)?.attrs.styleId).toBe('Bearbeitet')
      expect(post).not.toHaveBeenCalled()
    })

    it('still takes an edit of a field that holds text', () => {
      const post = vi.fn<(msg: EditorToHost) => void>()
      const state = EditorState.create({
        schema,
        doc: doc(para(0, 'Name'), fieldPara(1, 'Muster')),
        plugins: [markersPlugin, editScopePlugin(post)],
      })
      view = new EditorView(document.createElement('div'), { state })
      setMarkers(view, [field(1)])
      setEditScope('everything')
      const pos = start(view, 1)
      const node = view.state.doc.nodeAt(pos)!
      view.dispatch(
        view.state.tr.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          fieldDisplay: { ...node.attrs.fieldDisplay, left: 'Beispiel' },
        }),
      )
      expect(view.state.doc.nodeAt(pos)?.attrs.fieldDisplay.left).toBe('Beispiel')
      expect(post).not.toHaveBeenCalled()
    })
  })
})
