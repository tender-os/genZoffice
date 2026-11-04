// The host's slot markers inside the editor: a ProseMirror plugin that keeps them as node decorations
// across edits, phased loading and re-parses, plus the calls the shim makes for vound:markers,
// vound:scrollTo and vound:focus. Decorations live in the view only, so nothing here reaches a save.
// In review mode the chip (fieldChip.ts) marks the hovered block, and a save does not seal the markers: they
// stay until the app's re-parse of the saved bytes replaces the document, and then only if the host posted
// them again after the save (it does once it has checked the saved file's numbering).
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import { ReplaceStep } from '@tiptap/pm/transform'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { SlotMarker } from './embed-protocol'
import { HOT_CLASS, hasSourceIndices, markerDecorations } from './markers'

/** `placedOnce`: a build since the last vound:markers ran against a parsed document (one whose nodes carry
 *  source indices), whether or not any marker matched; until then an edit may still bring the document.
 *  `awaitingDocument`: the document of the last vound:open (or the editor's first one) has not replaced the
 *  content yet; its whole-document replace rebuilds the markers, any later one clears them.
 *  `hot`: the start of the marked block under the pointer (fieldChip.ts), mapped through edits.
 *  `afterSave`: a review-mode save happened and the app may still replace the document with the re-parsed
 *  saved bytes ('awaitingReparse'); 'reconfirmed' once the host posted markers after that save */
export type MarkersState = {
  markers: SlotMarker[]
  current: string | null
  sealed: boolean
  placedOnce: boolean
  awaitingDocument: boolean
  afterSave: 'idle' | 'awaitingReparse' | 'reconfirmed'
  hot: number | null
  set: DecorationSet
}
type MarkersMeta = {
  markers?: SlotMarker[]
  current?: string | null
  sealed?: true
  reset?: true
  saved?: true
  hot?: number | null
}

export const markersKey = new PluginKey<MarkersState>('voundMarkers')

const initial: MarkersState = {
  markers: [],
  current: null,
  sealed: false,
  placedOnce: false,
  awaitingDocument: true,
  afterSave: 'idle',
  hot: null,
  set: DecorationSet.empty,
}

function rebuild(doc: PmNode, state: MarkersState): MarkersState {
  const { decorations, dropped } = markerDecorations(doc, state.markers, state.current)
  if (dropped.length > 0) console.info('[vound] markers dropped', dropped)
  const set = DecorationSet.create(doc, decorations)
  const placedOnce = state.placedOnce || decorations.length > 0 || (state.markers.length > 0 && hasSourceIndices(doc))
  return { ...state, placedOnce, set }
}

/** a step that replaced the whole document, as Tiptap's setContent does when a document opens */
export const replacedWholeDocument = (tr: Transaction) =>
  tr.steps.some((step, index) => step instanceof ReplaceStep && step.from === 0 && step.to === tr.docs[index].content.size)

/** a document change: map the decorations, and walk the document again only when mapping cannot be trusted */
function afterDocChange(tr: Transaction, old: MarkersState): MarkersState {
  if (replacedWholeDocument(tr)) {
    // the document of the last open: markers built against the previous document while the parse ran
    // have nothing to map from, so they are built again on the new one
    if (old.awaitingDocument) return rebuild(tr.doc, { ...old, awaitingDocument: false })
    // the re-parse after a review-mode save: the host's markers posted since the save name the saved file,
    // so they are built on it; older ones may name paragraphs the save renumbered, so they go
    if (old.afterSave === 'reconfirmed') return rebuild(tr.doc, { ...old, afterSave: 'idle', hot: null })
    if (old.afterSave === 'awaitingReparse') return { ...initial, awaitingDocument: false, current: old.current }
    // a document that came without vound:open (File > Open inside the editor, a new blank document) or
    // replaced everything afterwards: the host's indices name another document, so the markers go
    return { ...initial, awaitingDocument: false }
  }
  const hot = old.hot === null ? null : tr.mapping.map(old.hot, 1)
  // the app replaces the document after a save only while it is still the one it saved: after any other
  // change no re-parse follows
  const kept: MarkersState = { ...old, hot, afterSave: 'idle' }
  const mapped = old.set.map(tr.mapping, tr.doc)
  const oldCount = old.set.find().length
  // phased appends, splits and deletions change the block count; a split leaves one index on two nodes
  const blockCountChanged = tr.before.childCount !== tr.doc.childCount
  // setNodeMarkup (updateAttributes, every attribute command) replaces the node and drops its node
  // decoration, while docxPara rides along in the copied attrs
  const decorationDropped = mapped.find().length < oldCount
  // markers that arrived before their document: the opened document replaced the initial one-block
  // content without changing the block count. Once a build saw a parsed document this stays off, so
  // markers that all missed it do not cost a walk per keystroke
  const neverPlaced = oldCount === 0 && !old.placedOnce
  if (blockCountChanged || decorationDropped || neverPlaced) return rebuild(tr.doc, kept)
  return { ...kept, set: mapped }
}

function apply(tr: Transaction, old: MarkersState): MarkersState {
  const meta: MarkersMeta | undefined = tr.getMeta(markersKey)
  if (meta?.sealed) return { ...initial, sealed: true }
  if (meta?.reset) return initial
  if (old.sealed) return old
  // a save in review mode: the markers stay on the nodes they mark; see afterDocChange for the re-parse
  if (meta?.saved) return { ...old, afterSave: 'awaitingReparse' }
  if (meta && meta.hot !== undefined) return { ...old, hot: meta.hot }
  if (meta?.markers) {
    const afterSave = old.afterSave === 'idle' ? 'idle' : 'reconfirmed'
    return rebuild(tr.doc, { ...old, markers: meta.markers, placedOnce: false, afterSave })
  }
  if (meta && meta.current !== undefined) return rebuild(tr.doc, { ...old, current: meta.current })
  if (tr.docChanged) return afterDocChange(tr, old)
  return old
}

export const markersPlugin = new Plugin<MarkersState>({
  key: markersKey,
  state: { init: () => initial, apply },
  props: {
    decorations: (state) => viewDecorations(state),
  },
})

/** what the view draws: the marker decorations, plus the hover outline on the hot block while a marker
 *  decoration still starts there */
export function viewDecorations(state: EditorState) {
  const markers = markersKey.getState(state)
  if (!markers || markers.hot === null) return markers?.set ?? DecorationSet.empty
  const hot = markers.hot
  const block = markers.set.find(hot, hot).find((decoration) => decoration.from === hot)
  if (!block) return markers.set
  return markers.set.add(state.doc, [Decoration.node(block.from, block.to, { class: HOT_CLASS }, { hot: true })])
}

const dispatchMeta = (view: EditorView, meta: MarkersMeta) => view.dispatch(view.state.tr.setMeta(markersKey, meta))

/** replace all markers (vound:markers) */
export function setMarkers(view: EditorView, markers: SlotMarker[]) {
  dispatchMeta(view, { markers })
}

/** make one marker current, or none (vound:focus) */
export function setCurrent(view: EditorView, id: string | null) {
  dispatchMeta(view, { current: id })
}

/** after an explicit save outside review mode: drop every marker and ignore markers and focus until the
 *  next document opens. In review mode: keep the markers until the app's re-parse replaces the document,
 *  and build on it only the markers the host posted after the save */
export const saveMeta = (mode: 'review' | 'plain'): MarkersMeta => (mode === 'review' ? { saved: true } : { sealed: true })

/** an explicit save (embeddedSave.ts), in review mode or not */
export function markersOnSave(view: EditorView, mode: 'review' | 'plain') {
  dispatchMeta(view, saveMeta(mode))
}

/** mark the block starting at `pos` as hovered, or none (fieldChip.ts) */
export function setHot(view: EditorView, pos: number | null) {
  if (markersKey.getState(view.state)?.hot === pos) return
  dispatchMeta(view, { hot: pos })
}

/** a document opens (vound:open): drop the previous document's markers and the seal, so no old index
 *  is rebuilt onto the new document's paragraphs; markers the host posts after the open apply as usual */
export function resetMarkers(view: EditorView) {
  dispatchMeta(view, { reset: true })
}

/** scroll to a node, gliding there unless the viewer asked for reduced motion */
export function glideIntoView(dom: HTMLElement, block: ScrollLogicalPosition) {
  const reduced = window.matchMedia?.('(prefers-reduced-motion: reduce)').matches
  dom.scrollIntoView({ block, behavior: reduced ? 'auto' : 'smooth' })
}

/** make the marker current and scroll its node to the middle of the view; false when no decoration carries
 *  the id (a node shared by several slots carries all of them) */
export function scrollToMarker(view: EditorView, id: string) {
  const set = markersKey.getState(view.state)?.set
  const [decoration] = set?.find(undefined, undefined, (spec) => spec.slotIds.includes(id)) ?? []
  if (!decoration) return false
  setCurrent(view, id)
  const dom = view.nodeDOM(decoration.from)
  if (dom instanceof HTMLElement) glideIntoView(dom, 'center')
  return true
}
