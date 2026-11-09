// The field chip in review mode: hovering a marked block (a paragraph or table cell carrying a slot marker)
// shows a small pill in the empty gutter right of the page, vertically centred on the block, with the
// buttons the host sent for each of the block's fields (SlotMarker.actions). A dashed tether runs at the
// block's height across the page's right margin to the chip, starting at the field's ring; in a cell with
// another cell after it in its row the tether leaves the ring's bottom right corner, runs along its bottom
// edge past the table and bends up into the margin, so it covers no text and no neighbour owns it; in a
// multi-column section the tether is left out, the next column may lie in its way. A block holding several
// fields gets one labelled row per field, in the order of its data-vound-slot ids. A press posts
// vound:fieldAction; the host decides what it does.
//
// The chip and tether live in a zero-height layer inside the page's scroll container, so they scroll with
// the document; a scroll or a resize hides them anyway. They stay while the pointer moves from the block to
// the chip and hide 220ms after it left both. The chip never grows past the gutter (reviewGutter.ts, which
// also keeps that gutter free): labels shrink or wrap above their buttons, and only a gutter narrower than
// the buttons hides the chip (the host's bar offers the current field's buttons then); the hover outline
// stays. The hovered block is a decoration (markersPlugin.ts, setHot), which fieldRing.ts rings: a class
// written onto the block's DOM would read as a DOM change to ProseMirror.
import { embedLang } from './embedLocale'
import { fillWithAiLabel } from './fillWithAi'
import type { EditorView } from '@tiptap/pm/view'
import type { EditorToHost, FieldAction, MarkerTone, SlotMarker } from './embed-protocol'
import { MARKER_CLASS } from './markers'
import { markersKey, setHot } from './markersPlugin'
import { ringBox, runAround } from './fieldRing'
import { chipRoom, chipRowWidth, PAGE_GAP_PX } from './reviewGutter'

const HIDE_DELAY_MS = 220

const BUTTONS: Record<FieldAction, { icon: string; title: string }> = {
  accept: { icon: '✓', title: 'Übernehmen' },
  reject: { icon: '✕', title: 'Ablehnen' },
  fill: { icon: '✎', title: 'Wert eintragen' },
  ask: { icon: '✦', title: 'Agent fragen' },
  restore: { icon: '↺', title: 'KI-Wert wiederherstellen' },
}

/** tones of a field with no value yet: their tether is amber */
const MISSING_TONES: MarkerTone[] = ['gap', 'conditional', 'refused', 'otherParty']

/** the page's scroll container (App.tsx's main.editor-scroll); the editor root's parent without one */
const scrollContainer = (view: EditorView) => view.dom.closest('.editor-scroll') ?? view.dom.parentElement

/** the innermost marked block of this document around an element */
function markedBlock(view: EditorView, target: Element) {
  const block = target.closest(`.${MARKER_CLASS}`)
  return block && view.dom.contains(block) ? block : null
}

const slotIdsOf = (block: Element) => (block.getAttribute('data-vound-slot') ?? '').split(' ').filter((id) => id !== '')

const markersOf = (view: EditorView) => markersKey.getState(view.state)?.markers ?? []

/** the markers of a block's fields that offer at least one button, in the block's id order */
function fieldsWithButtons(view: EditorView, ids: string[]) {
  const markers = markersOf(view)
  return ids
    .map((id) => markers.find((marker) => marker.id === id))
    .filter((marker): marker is SlotMarker => marker !== undefined && marker.actions.length > 0)
}

/** the editor root's child holding the block: the paragraph itself, or the table around a cell */
function pageLevelBlock(view: EditorView, block: Element) {
  let element = block
  while (element.parentElement && element.parentElement !== view.dom) element = element.parentElement
  return element
}

/** whether the block sits in a column of a multi-column section: another column may lie right of it */
function inColumns(view: EditorView, block: Element) {
  if (block.closest('.doc-col-block')) return true
  const columns = Number.parseInt(view.dom.ownerDocument.defaultView?.getComputedStyle(view.dom).columnCount ?? '', 10)
  return columns > 1
}

/** the narrowest chip for these fields: the row with the most buttons (labels shrink or wrap) */
const narrowestChip = (fields: SlotMarker[]) => Math.max(...fields.map((marker) => chipRowWidth(marker.actions.length)))

/** whether the block's cell, or a cell around it, has another cell after it in its row */
function cellBeside(view: EditorView, block: Element) {
  for (let cell = block.closest('td, th'); cell && view.dom.contains(cell); cell = cell.parentElement?.closest('td, th') ?? null) {
    if (cell.nextElementSibling) return true
  }
  return false
}

/** where the tether bends up from the field's ring bottom: halfway across the margin; null with no cell beside */
function elbowBend(view: EditorView, block: Element, contentRight: number, chipLeft: number) {
  if (!cellBeside(view, block)) return null
  return (contentRight + chipLeft) / 2
}

/** the document position the block's marker decoration starts at */
const SPARKLES_SVG =
  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
  '<path d="M9.937 15.5A2 2 0 0 0 8.5 14.063l-6.135-1.582a.5.5 0 0 1 0-.962L8.5 9.936A2 2 0 0 0 9.937 8.5l1.582-6.135a.5.5 0 0 1 .963 0L14.063 8.5A2 2 0 0 0 15.5 9.937l6.135 1.581a.5.5 0 0 1 0 .964L15.5 14.063a2 2 0 0 0-1.437 1.437l-1.582 6.135a.5.5 0 0 1-.963 0z"/>' +
  '<path d="M20 3v4M22 5h-4M4 17v2M5 18H3"/></svg>'

function blockStart(view: EditorView, block: Element) {
  const decorations = markersKey.getState(view.state)?.set.find() ?? []
  return decorations.find((decoration) => view.nodeDOM(decoration.from) === block)?.from ?? null
}

function chipButton(marker: SlotMarker, action: FieldAction) {
  const button = document.createElement('button')
  button.type = 'button'
  button.className = `vound-chip-btn vound-chip-${action}`
  if (action === 'ask') {
    const label = fillWithAiLabel(embedLang(location.search, true))
    button.innerHTML = SPARKLES_SVG
    button.dataset.label = label
    button.setAttribute('aria-label', label)
  } else {
    button.textContent = BUTTONS[action].icon
    button.title = BUTTONS[action].title
    button.setAttribute('aria-label', BUTTONS[action].title)
  }
  button.dataset.id = marker.id
  button.dataset.action = action
  return button
}

function chipRow(marker: SlotMarker, labelled: boolean) {
  const row = document.createElement('div')
  row.className = 'vound-chip-row'
  row.setAttribute('role', 'group')
  row.setAttribute('aria-label', marker.label)
  if (labelled) {
    const label = document.createElement('span')
    label.className = 'vound-chip-label'
    label.textContent = marker.label
    // the label may be cut short in a narrow gutter
    label.title = marker.label
    row.append(label)
  }
  row.append(...marker.actions.map((action) => chipButton(marker, action)))
  return row
}

/** Shows the chip for hovered marked blocks of `view` and posts a pressed button through `post`; returns
 *  the unmount. Mounted for review mode only (reviewMode.ts). */
export function mountFieldChip(view: EditorView, post: (msg: EditorToHost) => void): () => void {
  const ownerDocument = view.dom.ownerDocument
  const layer = ownerDocument.createElement('div')
  layer.className = 'vound-chip-layer'
  const tether = ownerDocument.createElement('div')
  tether.className = 'vound-chip-tether'
  tether.hidden = true
  const elbow = ownerDocument.createElement('div')
  elbow.className = 'vound-chip-tether vound-chip-tether-elbow'
  elbow.hidden = true
  const chip = ownerDocument.createElement('div')
  chip.className = 'vound-chip'
  chip.hidden = true
  layer.append(tether, elbow, chip)

  let hoveredBlock: Element | null = null
  let hideTimer: ReturnType<typeof setTimeout> | undefined

  const cancelHide = () => clearTimeout(hideTimer)
  const hideChip = () => {
    chip.hidden = true
    tether.hidden = true
    elbow.hidden = true
  }
  const hide = () => {
    cancelHide()
    hideChip()
    hoveredBlock = null
    setHot(view, null)
  }
  const scheduleHide = () => {
    cancelHide()
    hideTimer = setTimeout(hide, HIDE_DELAY_MS)
  }

  /** the scroll container, with the layer in it (the app may have re-rendered the container) */
  const placeLayer = () => {
    const container = scrollContainer(view)
    if (container && layer.parentElement !== container) container.append(layer)
    return container
  }

  /** fill the chip, then show it beside the block if the gutter holds it; labels shrink or wrap above
   *  their buttons to stay inside the gutter */
  const showChip = (block: Element, fields: SlotMarker[], labelled: boolean) => {
    const container = placeLayer()
    if (!container) return hideChip()
    const page = view.dom.getBoundingClientRect()
    const box = block.getBoundingClientRect()
    const contentRight = Math.max(box.right, pageLevelBlock(view, block).getBoundingClientRect().right)
    const bounds = container.getBoundingClientRect()
    // a block wider than the page (a table running into the margin) moves the chip further out
    const anchor = Math.max(page.right, contentRight)
    const scrollbar = container instanceof HTMLElement ? container.offsetWidth - container.clientWidth : 0
    const room = chipRoom(bounds.right - scrollbar - anchor)
    if (room < narrowestChip(fields)) return hideChip()
    chip.replaceChildren(...fields.map((marker) => chipRow(marker, labelled)))
    chip.classList.toggle('vound-chip-rows', fields.length > 1)
    chip.style.maxWidth = `${room}px`
    chip.hidden = false
    const origin = layer.getBoundingClientRect()
    const chipLeft = anchor + PAGE_GAP_PX
    const middle = box.top + box.height / 2 - origin.top
    const end = chipLeft - origin.left
    const ring = ringBox(runAround(block), origin)
    const fieldRight = ring.left + ring.width
    const start = contentRight - origin.left
    const bend = elbowBend(view, block, start, end)
    const tetherLeft = bend ?? Math.max(fieldRight, start)
    chip.style.left = `${end}px`
    chip.style.top = `${middle}px`
    tether.style.left = `${tetherLeft}px`
    tether.style.width = `${end - tetherLeft}px`
    tether.style.top = `${middle}px`
    elbow.style.left = `${fieldRight}px`
    elbow.style.width = `${tetherLeft - fieldRight}px`
    elbow.style.top = `${middle}px`
    elbow.style.height = `${ring.top + ring.height - middle}px`
    const missing = MISSING_TONES.some((tone) => block.classList.contains(`${MARKER_CLASS}-${tone}`))
    tether.classList.toggle('vound-chip-tether-missing', missing)
    elbow.classList.toggle('vound-chip-tether-missing', missing)
    // in a column the next column may lie between the block and the margin: the chip shows, the tether not
    tether.hidden = inColumns(view, block)
    elbow.hidden = tether.hidden || bend === null
  }

  const show = (block: Element) => {
    cancelHide()
    if (block === hoveredBlock) return
    hoveredBlock = block
    setHot(view, blockStart(view, block))
    const ids = slotIdsOf(block)
    const fields = fieldsWithButtons(view, ids)
    if (fields.length === 0) return hideChip()
    showChip(block, fields, ids.length > 1)
  }

  const onOver = (event: MouseEvent) => {
    if (!(event.target instanceof Element)) return
    if (chip.contains(event.target)) return cancelHide()
    const block = markedBlock(view, event.target)
    if (block) show(block)
    else if (hoveredBlock) scheduleHide()
  }
  // the pointer left the document (relatedTarget null): mouseover fires nowhere else
  const onOut = (event: MouseEvent) => {
    if (event.relatedTarget === null && hoveredBlock) scheduleHide()
  }
  // a press keeps the focus, and with it the caret, in the document
  const keepFocus = (event: MouseEvent) => event.preventDefault()
  const onClick = (event: MouseEvent) => {
    const button = event.target instanceof Element ? event.target.closest('button') : null
    if (!button) return
    // the markers may have changed since the chip was filled: post only what the field still offers
    const marker = markersOf(view).find((candidate) => candidate.id === button.dataset.id)
    const action = marker?.actions.find((candidate) => candidate === button.dataset.action)
    hide()
    if (marker && action) post({ type: 'vound:fieldAction', id: marker.id, action })
  }

  ownerDocument.addEventListener('mouseover', onOver)
  ownerDocument.addEventListener('mouseout', onOut)
  // scroll does not bubble; a capture listener on the document hears the scroll container's
  ownerDocument.addEventListener('scroll', hide, true)
  // a resize (or the app's zoom refit that follows it) moves the page under a placed chip
  ownerDocument.defaultView?.addEventListener('resize', hide)
  chip.addEventListener('mousedown', keepFocus)
  chip.addEventListener('click', onClick)

  return () => {
    hide()
    ownerDocument.removeEventListener('mouseover', onOver)
    ownerDocument.removeEventListener('mouseout', onOut)
    ownerDocument.removeEventListener('scroll', hide, true)
    ownerDocument.defaultView?.removeEventListener('resize', hide)
    layer.remove()
  }
}
