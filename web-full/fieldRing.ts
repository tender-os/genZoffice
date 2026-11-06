// The selected and the hover ring of review fields. A field can span consecutive blocks (paragraphs, list
// items, or the cells of one column in consecutive rows); its ring is one closed box around the whole run,
// so no line runs between its blocks and none goes missing at the run's edge. The rings live in a
// zero-size layer at the end of the page's scroll container, above the page: a ring drawn on a block would
// sit under the tint of the marked block after it, which paints later. The layer takes no part in layout.
import { Plugin, PluginKey } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { CURRENT_CLASS, HOT_CLASS, MARKER_CLASS } from './markers'

type Ring = { kind: 'current' | 'hot'; left: number; top: number; width: number; height: number; tone: Element }

const sameFields = (a: Element, b: Element | null | undefined): b is Element =>
  !!b && b.classList.contains(MARKER_CLASS) && b.getAttribute('data-vound-slot') === a.getAttribute('data-vound-slot')

/** the cell of the next (step 1) or previous (step -1) row in the same column */
function cellBeside(cell: HTMLTableCellElement, step: 1 | -1) {
  const row = step === 1 ? cell.parentElement?.nextElementSibling : cell.parentElement?.previousElementSibling
  return row instanceof HTMLTableRowElement ? row.cells[cell.cellIndex] : undefined
}

/** the block after (1) or before (-1) this one in its run: the adjacent sibling, or the cell in the
 *  adjacent row, when it carries the same fields */
function neighbourInRun(block: Element, step: 1 | -1) {
  if (block instanceof HTMLTableCellElement) {
    const cell = cellBeside(block, step)
    return sameFields(block, cell) && cell.colSpan === block.colSpan ? cell : null
  }
  const sibling = step === 1 ? block.nextElementSibling : block.previousElementSibling
  return sameFields(block, sibling) ? sibling : null
}

/** the whole run around a block */
export function runAround(block: Element): Element[] {
  const run = [block]
  for (let prev = neighbourInRun(block, -1); prev; prev = neighbourInRun(prev, -1)) run.unshift(prev)
  for (let next = neighbourInRun(block, 1); next; next = neighbourInRun(next, 1)) run.push(next)
  return run
}

/** blocks in document order, split into runs */
export function runsOf(blocks: Element[]): Element[][] {
  const runs: Element[][] = []
  for (const block of blocks) {
    const last = runs.at(-1)
    if (last && neighbourInRun(last[last.length - 1], 1) === block) last.push(block)
    else runs.push([block])
  }
  return runs
}

/** how far a block's tint reaches past its box (host.css --vs-l, --vs-r, --vs-y) */
function reach(block: Element) {
  const style = block.ownerDocument.defaultView?.getComputedStyle(block)
  const px = (name: string) => Number.parseFloat(style?.getPropertyValue(name) ?? '') || 0
  return { left: px('--vs-l'), right: px('--vs-r'), y: px('--vs-y') }
}

/** the box a run's ring goes around, relative to `origin` */
export function ringBox(run: Element[], origin: { left: number; top: number }) {
  const rects = run.map((block) => block.getBoundingClientRect())
  const out = reach(run[0])
  const left = Math.min(...rects.map((rect) => rect.left)) - out.left
  const right = Math.max(...rects.map((rect) => rect.right)) + out.right
  const top = Math.min(...rects.map((rect) => rect.top)) - out.y
  const bottom = Math.max(...rects.map((rect) => rect.bottom)) + out.y
  return { left: left - origin.left, top: top - origin.top, width: right - left, height: bottom - top }
}

/** the rings to draw: one per run of selected blocks, and one around the hovered block's run unless selected */
export function ringsIn(root: Element, origin: { left: number; top: number }): Ring[] {
  const current = runsOf([...root.querySelectorAll(`.${CURRENT_CLASS}`)])
  const rings: Ring[] = current.map((run) => ({ kind: 'current', tone: run[0], ...ringBox(run, origin) }))
  const hot = root.querySelector(`.${HOT_CLASS}:not(.${CURRENT_CLASS})`)
  if (hot) {
    const run = runAround(hot)
    rings.push({ kind: 'hot', tone: hot, ...ringBox(run, origin) })
  }
  return rings
}

/** the page's scroll container (App.tsx's main.editor-scroll); the editor root's parent without one */
const scrollContainer = (view: EditorView) => view.dom.closest('.editor-scroll') ?? view.dom.parentElement

function ringElement(ring: Ring, ownerDocument: Document) {
  const element = ownerDocument.createElement('div')
  element.className = `vound-ring vound-ring-${ring.kind}`
  const style = ownerDocument.defaultView?.getComputedStyle(ring.tone)
  element.style.setProperty('--vr-ring', style?.getPropertyValue(ring.kind === 'hot' ? '--vs-hover' : '--vs-ring') ?? '')
  element.style.setProperty('--vr-halo', style?.getPropertyValue('--vs-halo') ?? '')
  element.style.left = `${ring.left}px`
  element.style.top = `${ring.top}px`
  element.style.width = `${ring.width}px`
  element.style.height = `${ring.height}px`
  return element
}

/** the rings of the marked blocks under `root`, drawn into `layer` */
export function drawRings(root: Element, layer: HTMLElement) {
  const origin = layer.getBoundingClientRect()
  layer.replaceChildren(...ringsIn(root, origin).map((ring) => ringElement(ring, layer.ownerDocument)))
}

/** draws the rings of `view` and keeps them on their blocks; returns the unmount */
function mountRings(view: EditorView) {
  const ownerDocument = view.dom.ownerDocument
  const layer = ownerDocument.createElement('div')
  layer.className = 'vound-ring-layer'
  const draw = () => {
    const container = scrollContainer(view)
    if (!container) return
    if (layer.parentElement !== container) container.append(layer)
    drawRings(view.dom, layer)
  }
  // a reflow (pagination, zoom, fonts, images) moves blocks without a state change
  const observer = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(draw)
  observer?.observe(view.dom)
  ownerDocument.defaultView?.addEventListener('resize', draw)
  draw()
  return {
    draw,
    unmount: () => {
      observer?.disconnect()
      ownerDocument.defaultView?.removeEventListener('resize', draw)
      layer.remove()
    },
  }
}

export const fieldRingKey = new PluginKey('voundFieldRing')

/** draws the rings after every update of the editor (review mode, reviewMode.ts) */
export const fieldRingPlugin = new Plugin({
  key: fieldRingKey,
  view: (view) => {
    const rings = mountRings(view)
    return { update: rings.draw, destroy: rings.unmount }
  },
})
