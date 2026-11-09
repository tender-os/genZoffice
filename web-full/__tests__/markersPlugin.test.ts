import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { EditorState, type Transaction } from '@tiptap/pm/state'
import { DecorationSet } from '@tiptap/pm/view'
import type { SlotMarker } from '../embed-protocol'
import { markersKey, markersPlugin, saveMeta, viewDecorations } from '../markersPlugin'
import { cell, doc, para, paragraphs, schema, table } from './markerSchema'
import { splitSourceParagraphPlugin } from '../../apps/docs/src/renderer/editor/split-source-paragraph'
import { locateBlock } from '../markers'

const marker = (id: string, paragraph: number): SlotMarker => ({ id, paragraph, scope: 'paragraph', tone: 'proposed', label: id, actions: [] })
const stateWith = (...blocks: Parameters<typeof doc>) => EditorState.create({ schema, doc: doc(...blocks), plugins: [markersPlugin] })
/** the initial content the editor mounts with before a document is opened */
const initialState = () => stateWith(para(null))
const apply = (state: EditorState, edit: (tr: Transaction) => Transaction) => state.apply(edit(state.tr))
const withMeta = (state: EditorState, meta: object) => apply(state, (tr) => tr.setMeta(markersKey, meta))
const decorated = (state: EditorState) =>
  markersKey.getState(state)!.set.find().map((d) => ({ id: d.spec.slotIds.join(' '), from: d.from, to: d.to }))
const replaceDocument = (state: EditorState, blocks: ReturnType<typeof para>[]) =>
  apply(state, (tr) => tr.replaceWith(0, tr.doc.content.size, blocks))

let info: ReturnType<typeof vi.spyOn>
beforeEach(() => { info = vi.spyOn(console, 'info').mockImplementation(() => {}) })
afterEach(() => info.mockRestore())

describe('markersPlugin', () => {
  it('applies markers that arrived before the document and after phased appends', () => {
    let state = withMeta(initialState(), { markers: [marker('s0', 0), marker('s10', 10), marker('s190', 190)] })
    expect(decorated(state)).toEqual([])

    // phase 1: the first 64 blocks replace the initial paragraph
    state = replaceDocument(state, paragraphs(0, 64))
    expect(decorated(state).map((d) => d.id)).toEqual(['s0', 's10'])

    // the tail lands in one append
    state = apply(state, (tr) => tr.insert(tr.doc.content.size, paragraphs(64, 200)))
    expect(state.doc.childCount).toBe(200)
    const late = decorated(state).find((d) => d.id === 's190')
    expect(late).toBeDefined()
    expect(state.doc.nodeAt(late!.from)?.attrs.docxPara).toBe(190)
  })

  it('maps decorations through a text edit, marks both halves of a split paragraph and drops a copied index', () => {
    let state = withMeta(stateWith(...paragraphs(0, 6)), { markers: [marker('s4', 4)] })
    const [before] = decorated(state)
    expect(state.doc.nodeAt(before.from)?.attrs.docxPara).toBe(4)

    // typing in the first paragraph moves the decoration
    state = apply(state, (tr) => tr.insertText('xx', 1))
    expect(decorated(state)).toEqual([{ id: 's4', from: before.from + 2, to: before.to + 2 }])

    // a split copies docxPara 4 to both halves, side by side: the second continues the first
    state = apply(state, (tr) => tr.split(decorated(state)[0].from + 4))
    expect(state.doc.childCount).toBe(7)
    expect(decorated(state).map((d) => [d.id, state.doc.nodeAt(d.from)?.textContent])).toEqual([['s4', 'par'], ['s4', 'agraph 4']])
    expect(info).not.toHaveBeenCalled()

    // the same index on a paragraph elsewhere names two places: the marker goes
    state = apply(state, (tr) => tr.insert(tr.doc.content.size, para(4, 'copy')))
    expect(decorated(state)).toEqual([])
    expect(info).toHaveBeenCalledWith('[vound] markers dropped', ['s4'])
  })

  it('survives an attribute change and a one-block replacement', () => {
    let state = withMeta(stateWith(...paragraphs(0, 3)), { markers: [marker('s1', 1)] })
    const [before] = decorated(state)
    state = apply(state, (tr) => tr.setNodeMarkup(before.from, undefined, { docxPara: 1, align: 'center' }))
    expect(state.doc.nodeAt(before.from)?.attrs.align).toBe('center')
    expect(decorated(state)).toEqual([before])

    // markers pending on the initial one-paragraph content, then a one-paragraph document opens
    let single = withMeta(initialState(), { markers: [marker('s0', 0)] })
    expect(decorated(single)).toEqual([])
    single = replaceDocument(single, [para(0, 'only paragraph')])
    expect(decorated(single)).toEqual([{ id: 's0', from: 0, to: single.doc.child(0).nodeSize }])
  })

  it('does not walk the document per edit once its markers were placed and then all dropped', () => {
    let state = withMeta(stateWith(...paragraphs(0, 3)), { markers: [marker('s1', 1)] })
    state = apply(state, (tr) => tr.delete(decorated(state)[0].from, decorated(state)[0].to))
    expect(decorated(state)).toEqual([])
    info.mockClear()
    state = apply(state, (tr) => tr.insertText('typed', 1))
    expect(info).not.toHaveBeenCalled()
    expect(decorated(state)).toEqual([])
  })

  it('does not walk the document per edit when every marker missed a stamped document', () => {
    // all indices stale or in nested tables: nothing is placed, but the document carries source indices
    const state = withMeta(stateWith(...paragraphs(0, 3)), { markers: [marker('stale', 40), marker('nested', 41)] })
    expect(decorated(state)).toEqual([])
    expect(markersKey.getState(state)!.placedOnce).toBe(true)
    info.mockClear()
    apply(state, (tr) => tr.insertText('typed', 1))
    expect(info).not.toHaveBeenCalled()
  })

  it('keeps waiting while the document carries no source indices', () => {
    const state = withMeta(initialState(), { markers: [marker('s0', 0)] })
    expect(markersKey.getState(state)!.placedOnce).toBe(false)
  })

  it('builds the markers on the opened document when they arrived while the previous one was still shown', () => {
    let state = withMeta(stateWith(...paragraphs(0, 10)), { reset: true })
    // vound:markers lands before the app's async parse replaced the document: every index misses the old one
    state = withMeta(state, { markers: [marker('s50', 50)] })
    expect(decorated(state)).toEqual([])
    // the new document arrives in one setContent with the same block count
    state = replaceDocument(state, paragraphs(45, 55))
    expect(state.doc.childCount).toBe(10)
    expect(decorated(state).map((d) => d.id)).toEqual(['s50'])
  })

  it('sealed: markers and current metas are ignored until an open resets everything', () => {
    let state = withMeta(stateWith(...paragraphs(0, 3)), { markers: [marker('s0', 0), marker('s2', 2)] })
    state = withMeta(state, { current: 's2' })
    expect(markersKey.getState(state)!.current).toBe('s2')

    state = withMeta(state, { sealed: true })
    expect(markersKey.getState(state)).toMatchObject({ markers: [], current: null, sealed: true })
    expect(decorated(state)).toEqual([])

    state = withMeta(state, { markers: [marker('s1', 1)] })
    state = withMeta(state, { current: 's1' })
    state = apply(state, (tr) => tr.insertText('edit', 1))
    expect(markersKey.getState(state)).toMatchObject({ markers: [], current: null, sealed: true })
    expect(decorated(state)).toEqual([])

    // vound:open
    state = withMeta(state, { reset: true })
    expect(markersKey.getState(state)).toEqual({
      markers: [],
      current: null,
      sealed: false,
      placedOnce: false,
      awaitingDocument: true,
      afterSave: 'idle',
      hot: null,
      set: DecorationSet.empty,
    })
    state = withMeta(state, { markers: [marker('s1', 1)] })
    expect(decorated(state).map((d) => d.id)).toEqual(['s1'])
  })

  it('an open drops the previous document\'s markers before the new document lands', () => {
    let state = withMeta(stateWith(...paragraphs(0, 3)), { markers: [marker('old', 1)] })
    state = withMeta(state, { current: 'old' })
    expect(decorated(state).map((d) => d.id)).toEqual(['old'])

    state = withMeta(state, { reset: true })
    // the next document replaces the content with a different block count: nothing of the old markers comes back
    state = replaceDocument(state, paragraphs(0, 5))
    expect(decorated(state)).toEqual([])
    expect(markersKey.getState(state)).toMatchObject({ markers: [], current: null })

    // markers the host posts after the open still apply
    state = withMeta(state, { markers: [marker('new', 4)] })
    expect(decorated(state).map((d) => d.id)).toEqual(['new'])
  })

  it('clears the markers when a document replaces the opened one without vound:open', () => {
    let state = withMeta(stateWith(...paragraphs(0, 10)), { reset: true })
    state = withMeta(state, { markers: [marker('s50', 50)] })
    state = replaceDocument(state, paragraphs(45, 55))
    expect(decorated(state).map((d) => d.id)).toEqual(['s50'])

    // File > Open inside the editor (or a password-protected file through loadFile) brings a foreign
    // document that happens to carry index 50 too: the host's markers do not name it
    state = replaceDocument(state, paragraphs(45, 55))
    expect(decorated(state)).toEqual([])
    expect(markersKey.getState(state)).toMatchObject({ markers: [], current: null })

    // the next vound:open waits for its document again
    state = withMeta(state, { reset: true })
    state = withMeta(state, { markers: [marker('s46', 46)] })
    state = replaceDocument(state, paragraphs(40, 50))
    expect(decorated(state).map((d) => d.id)).toEqual(['s46'])
  })

  it('in review mode a save does not seal and later markers apply', () => {
    // the opened document has landed
    let state = replaceDocument(stateWith(para(null)), paragraphs(0, 5))
    state = withMeta(state, { markers: [marker('s1', 1), marker('s3', 3)] })
    state = withMeta(state, { current: 's3' })

    // the markers stay on their nodes through the save (a clean save is never re-parsed or re-sent)
    state = withMeta(state, saveMeta('review'))
    expect(markersKey.getState(state)).toMatchObject({ current: 's3', sealed: false, afterSave: 'awaitingReparse' })
    expect(decorated(state).map((d) => d.id)).toEqual(['s1', 's3'])

    // the host checked the saved file and posts its markers before the app's re-parse replaced the
    // document: they are built again on the re-parsed one
    state = withMeta(state, { markers: [marker('s1', 1), marker('s3', 3)] })
    state = replaceDocument(state, paragraphs(0, 5))
    expect(decorated(state).map((d) => d.id)).toEqual(['s1', 's3'])
    expect(markersKey.getState(state)).toMatchObject({ current: 's3', afterSave: 'idle' })

    // the re-parse comes first: markers from before the save may name renumbered paragraphs, so they go;
    // the host's markers posted afterwards apply
    state = withMeta(state, saveMeta('review'))
    state = replaceDocument(state, paragraphs(0, 6))
    expect(decorated(state)).toEqual([])
    expect(markersKey.getState(state)!.current).toBe('s3')
    state = withMeta(state, { markers: [marker('s2', 2)] })
    expect(decorated(state).map((d) => d.id)).toEqual(['s2'])

    // outside review mode a save still seals
    state = withMeta(state, saveMeta('plain'))
    state = withMeta(state, { markers: [marker('s2', 2)] })
    expect(markersKey.getState(state)!.sealed).toBe(true)
    expect(decorated(state)).toEqual([])
  })

  it('a re-parse that renumbered nothing keeps the markers until the host sends its own', () => {
    let state = replaceDocument(stateWith(para(null)), paragraphs(0, 5))
    state = withMeta(state, { markers: [marker('s1', 1), marker('s3', 3)] })
    state = withMeta(state, saveMeta('review'))
    state = replaceDocument(state, paragraphs(0, 5))
    expect(decorated(state).map((d) => d.id)).toEqual(['s1', 's3'])
    expect(markersKey.getState(state)!.afterSave).toBe('idle')
  })

  it('after a review-mode save that the app does not re-parse, a later whole-document replace clears the markers', () => {
    let state = replaceDocument(stateWith(para(null)), paragraphs(0, 5))
    state = withMeta(state, { markers: [marker('s1', 1)] })
    state = withMeta(state, saveMeta('review'))
    state = withMeta(state, { markers: [marker('s1', 1)] })
    // the reviewer typed on: the app re-parses only a document it still holds unchanged, so none follows
    state = apply(state, (tr) => tr.insertText('typed', 1))
    expect(markersKey.getState(state)!.afterSave).toBe('idle')
    expect(decorated(state).map((d) => d.id)).toEqual(['s1'])
    // a select-all that replaces everything is not that re-parse: the markers go, as without a save
    state = replaceDocument(state, paragraphs(0, 5))
    expect(decorated(state)).toEqual([])
  })

  it('the hover outline follows its block through edits and goes with the markers', () => {
    let state = withMeta(stateWith(...paragraphs(0, 3)), { markers: [marker('s1', 1)] })
    const [block] = decorated(state)
    state = withMeta(state, { hot: block.from })
    const hotClass = (s: EditorState) => viewDecorations(s).find(undefined, undefined, (spec) => spec.hot === true).map((d) => d.from)
    expect(hotClass(state)).toEqual([block.from])

    state = apply(state, (tr) => tr.insertText('xx', 1))
    expect(hotClass(state)).toEqual([block.from + 2])

    state = withMeta(state, { hot: null })
    expect(hotClass(state)).toEqual([])

    state = withMeta(state, { hot: decorated(state)[0].from })
    state = withMeta(state, { markers: [] })
    expect(hotClass(state)).toEqual([])
  })
  describe('an emptied field keeps its marker', () => {
    // a value cell with a marked paragraph, between two plain paragraphs
    const cellState = () =>
      withMeta(stateWith(para(0, 'before'), table(cell(para(1, 'Name')), cell(para(2, 'Muster GmbH'))), para(3, 'after')), {
        markers: [marker('s2', 2), { ...marker('c2', 2), scope: 'cell' }],
      })
    /** applies `edit` with the plugins' appended transactions, as the view dispatches it */
    const edit = (state: EditorState, change: (tr: Transaction) => Transaction) => state.applyTransaction(change(state.tr)).state
    const valueParagraph = (state: EditorState) => {
      let found: { pos: number; docxPara: unknown; text: string } | null = null
      state.doc.descendants((node, pos) => {
        if (node.isTextblock && node.textContent !== 'Name' && pos > 8 && pos < state.doc.content.size - 8) found = { pos, docxPara: node.attrs.docxPara, text: node.textContent }
        return !node.isTextblock
      })
      return found!
    }

    it('an edit that turns the emptied paragraph into a new default one', () => {
      let state = cellState()
      const { pos } = valueParagraph(state)
      state = edit(state, (tr) => tr.delete(pos + 1, pos + 1 + 'Muster GmbH'.length))
      state = edit(state, (tr) => tr.setNodeMarkup(pos, schema.nodes.docParagraph, null))
      expect(valueParagraph(state)).toEqual({ pos, docxPara: 2, text: '' })
      expect(decorated(state).map((d) => d.id).sort()).toEqual(['c2', 's2'])
    })

    it('a cut cell selection fills the cell with a new default paragraph', () => {
      let state = cellState()
      const { pos } = valueParagraph(state)
      const cellEnd = pos + state.doc.nodeAt(pos)!.nodeSize
      state = edit(state, (tr) => tr.replaceWith(pos, cellEnd, schema.nodes.docParagraph.create()))
      expect(valueParagraph(state)).toEqual({ pos, docxPara: 2, text: '' })
      expect(decorated(state).map((d) => d.id).sort()).toEqual(['c2', 's2'])
    })

    it('typing over a cell selection keeps the index on the paragraph that holds the new text', () => {
      let state = cellState()
      const { pos } = valueParagraph(state)
      const cellEnd = pos + state.doc.nodeAt(pos)!.nodeSize
      state = edit(state, (tr) => tr.replaceWith(pos, cellEnd, schema.nodes.docParagraph.create(null, schema.text('Beispiel AG'))))
      expect(valueParagraph(state)).toEqual({ pos, docxPara: 2, text: 'Beispiel AG' })
    })

    it('gives no index to a new paragraph while its old owner is still in the document', () => {
      let state = cellState()
      const { pos } = valueParagraph(state)
      state = edit(state, (tr) => tr.insert(pos, schema.nodes.docParagraph.create(null, schema.text('new'))))
      expect(state.doc.nodeAt(pos)!.attrs.docxPara).toBeNull()
      expect(state.doc.nodeAt(pos + 5)!.attrs.docxPara).toBe(2)
    })

    it('still drops the marker of a paragraph that was deleted outright', () => {
      let state = withMeta(stateWith(...paragraphs(0, 4)), { markers: [marker('s2', 2)] })
      const [block] = decorated(state)
      state = edit(state, (tr) => tr.delete(block.from, block.to))
      expect(decorated(state)).toEqual([])
      expect(info).toHaveBeenCalledWith('[vound] markers dropped', ['s2'])
    })
  })

  describe('a split field paragraph keeps its marker on both halves', () => {
    // the editor's own bookkeeping of a split source paragraph runs with the markers, as in the app
    const splitState = () =>
      withMeta(
        EditorState.create({ schema, doc: doc(para(0, 'Firma'), para(1, 'Muster GmbH'), para(2, 'Ort')), plugins: [splitSourceParagraphPlugin, markersPlugin] }),
        { markers: [marker('s1', 1)] },
      )
    const edit = (state: EditorState, change: (tr: Transaction) => Transaction) => state.applyTransaction(change(state.tr)).state
    const marked = (state: EditorState) => decorated(state).map((d) => [d.id, state.doc.nodeAt(d.from)!.textContent])
    const gmbh = (state: EditorState) => state.doc.child(1).content.size - 'GmbH'.length + state.doc.child(0).nodeSize + 1

    it('marks both halves after Enter, and a write or caret still finds the first one alone', () => {
      let state = edit(splitState(), (tr) => tr.split(gmbh(splitState())))
      expect(marked(state)).toEqual([['s1', 'Muster '], ['s1', 'GmbH']])
      expect(info).not.toHaveBeenCalled()
      expect(locateBlock(state.doc, { paragraph: 1, scope: 'paragraph' })?.node.textContent).toBe('Muster ')
      state = edit(state, (tr) => tr.insertText('Bau ', state.doc.child(0).nodeSize + 1 + 'Muster '.length))
      state = edit(state, (tr) => tr.insertText(' & Co. KG', state.doc.child(0).nodeSize + state.doc.child(1).nodeSize + 1 + 'GmbH'.length))
      expect(marked(state)).toEqual([['s1', 'Muster Bau '], ['s1', 'GmbH & Co. KG']])
    })

    it('gives the single-paragraph field back when the halves are joined', () => {
      let state = edit(splitState(), (tr) => tr.split(gmbh(splitState())))
      state = edit(state, (tr) => tr.join(state.doc.child(0).nodeSize + state.doc.child(1).nodeSize))
      expect(marked(state)).toEqual([['s1', 'Muster GmbH']])
      expect(state.doc.child(1).attrs).toMatchObject({ docxPara: 1, continuesPara: null })
    })
  })
})
