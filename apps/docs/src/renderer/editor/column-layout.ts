import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { EditorView } from '@tiptap/pm/view'
import { outerTableRows, splitFloatStyle } from '../pagination-lines'
import type { ColumnRowFragment, ColumnSplitShape, VerticalBlockSpec } from '../pagination-types'
import { TopLevelPositions } from './top-level-pos'
import { canvasZoomOf, type LayoutBatch } from './pagination-gaps'

const key = new PluginKey<DecorationSet>('columnLayout')

/**
 * Mixed-column canvas layout (documents whose sections disagree on column
 * count/width, so the uniform whole-page CSS multicol path can't apply).
 *
 * The engine keeps its 1-D single-flow virtual coordinates; this extension only
 * *paints* the flow into columns: each block in a multi-column region gets its
 * section's column width plus a constant per-column translate, and blocks of
 * later regions on the same page translate up over the space the columns
 * vacated (the page gap below pulls the next page up by the same amount).
 * Decorations are visual-only; measurement runs in the `.measuring-columns`
 * state, whose CSS neutralizes the transforms (widths stay — the engine's line
 * boxes must reflect column wrapping).
 *
 * A block is placed whole into the column containing its top, except a table
 * the engine cut at row boundaries: the rows of each later column ride their
 * own translate (doc-col-row), and a ghost clone after the table paints their
 * collapsed borders and any column-top header repeat (doc-col-host). Any other
 * block the engine split across columns paints overflowing its column bottom.
 */
export const ColumnLayoutExtension = Extension.create({
  name: 'columnLayout',

  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, set) {
            const next = tr.getMeta(key) as DecorationSet | undefined
            if (next) return next
            return set.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return key.getState(state)
          },
        },
      }),
    ]
  },
})

export interface ColumnBlockSpec {
  /** top-level block element (same anchoring as page gaps) */
  el: HTMLElement
  /** column width (px); absent for single-column regions (natural width) */
  widthPx?: number
  /** owning section's content width (--doc-content-w): tables resolve their spill/centering caps against it */
  contentWPx?: number
  /** owning section's side margins (--doc-margin-left/right overrides) */
  marginLeftPx?: number
  marginRightPx?: number
  /** owning section's top margin (--doc-margin-top override) in docs whose sections disagree on it */
  marginTopPx?: number
  /** owning section's typed docGrid pitch (pt) in mixed-grid docs; 0 = untyped
   *  section (opts out like snapToGrid=0 via doc-grid-nosnap) */
  gridPitchPt?: number
  /** owning section's character-grid letter-spacing delta (pt) in docs whose
   *  sections disagree on docGrid charSpace (doc-charspace-block) */
  charSpacePt?: number
  charSpaceSbPt?: number
  /** translate; dy < 0 pulls content up over vacated column space */
  dx: number
  dy: number
  /** left edge of the block's page on the shared paper (--page-cx); not undone by page-relative anchors */
  pageDx?: number
  /** vertical-text section block (doc-vert-block): writing-mode box, see verticalTextSpecs */
  vertical?: VerticalBlockSpec
  /** paragraph straddling two columns of different widths: a leading float
   *  (doc-col-split-float) rewraps the head or the tail at the narrower width */
  split?: ColumnSplitShape
  /** table cut at row boundaries into later columns (see columnLayoutSpecs) */
  rowFragments?: ColumnRowFragment[]
}

const SPLIT_CLASS = 'doc-col-split-float'

function splitFloatEl(style: string): HTMLElement {
  const span = document.createElement('span')
  span.className = SPLIT_CLASS
  span.contentEditable = 'false'
  span.setAttribute('aria-hidden', 'true')
  span.style.cssText = style
  return span
}

const PATCH_PROPS = [
  '--col-w',
  '--col-dx',
  '--col-dy',
  '--page-cx',
  '--doc-content-w',
  '--doc-margin-left',
  '--doc-margin-right',
  '--doc-margin-top',
  '--doc-grid-pitch',
  '--doc-char-space',
  '--doc-char-space-sb',
  '--vert-mb',
]

const PATCH_CLASSES = [
  'doc-col-block',
  'doc-grid-block',
  'doc-grid-nosnap',
  'doc-charspace-block',
  'doc-vert-block',
  'doc-vert-first',
  'doc-vert-tbRl',
  'doc-vert-tbRlV',
  'doc-vert-btLr',
]

/** pitch/charSpace-only specs get an inert marker class: .doc-col-block's
 *  width/transform rules must not touch blocks that only carry their section's
 *  grid pitch or character spacing. Pitch-0 opt-outs use their own class (same
 *  CSS as doc-nosnap): reusing doc-nosnap would trip sectionGridPitchSpecs'
 *  own-opt-out skip on the next pass and let patch cleanup strip a node's own
 *  snapToGrid=0 class. */
function specClass(spec: ColumnBlockSpec): string {
  const varsOnly =
    (spec.gridPitchPt !== undefined ||
      spec.charSpacePt !== undefined ||
      spec.marginTopPx !== undefined) &&
    spec.widthPx === undefined &&
    spec.contentWPx === undefined &&
    spec.dx === 0 &&
    spec.dy === 0 &&
    !spec.pageDx
  let cls = varsOnly ? 'doc-grid-block' : 'doc-col-block'
  if (spec.gridPitchPt === 0) cls += ' doc-grid-nosnap'
  if (spec.charSpacePt !== undefined) cls += ' doc-charspace-block'
  if (spec.vertical) {
    cls += ` doc-vert-block doc-vert-${spec.vertical.mode}`
    if (spec.vertical.firstOnPage) cls += ' doc-vert-first'
  }
  return cls
}

/** Rebuild the column-layout decorations (an empty list clears them). */
export function setColumnLayout(
  view: EditorView,
  specs: ColumnBlockSpec[],
  batch?: LayoutBatch,
): void {
  const decos: Decoration[] = []
  const positions = new TopLevelPositions(view)
  const styleOf = new Map<HTMLElement, { style: string; cls: string }>()
  const splitOf = new Map<HTMLElement, string>()
  const splitTables: Array<{ el: HTMLElement; fragments: ColumnRowFragment[] }> = []
  for (const spec of specs) {
    if (spec.split)
      splitOf.set(
        spec.el,
        splitFloatStyle(
          spec.split.floatPx,
          spec.split.heightPx,
          spec.split.insetPx,
          spec.split.rtl,
        ),
      )
    const style =
      (spec.widthPx !== undefined ? `--col-w:${round2(spec.widthPx)}px;` : '') +
      (spec.contentWPx !== undefined
        ? `--doc-content-w:${round2(spec.contentWPx)}px;--doc-margin-left:${round2(spec.marginLeftPx ?? 0)}px;--doc-margin-right:${round2(spec.marginRightPx ?? 0)}px;`
        : '') +
      (spec.marginTopPx !== undefined ? `--doc-margin-top:${round2(spec.marginTopPx)}px;` : '') +
      (spec.gridPitchPt !== undefined
        ? `--doc-grid-pitch:${spec.gridPitchPt > 0 ? `${round2(spec.gridPitchPt)}pt` : '0.0001px'};`
        : '') +
      (spec.charSpacePt !== undefined ? `--doc-char-space:${round4(spec.charSpacePt)}pt;` : '') +
      (spec.charSpaceSbPt !== undefined
        ? `--doc-char-space-sb:${round4(spec.charSpaceSbPt)}pt;`
        : '') +
      (spec.pageDx ? `--page-cx:${round2(spec.pageDx)}px;` : '') +
      (spec.vertical ? `--vert-mb:${round2(spec.vertical.marginBottomPx)}px;` : '') +
      `--col-dx:${round2(spec.dx)}px;--col-dy:${round2(spec.dy)}px`
    const cls = specClass(spec)
    styleOf.set(spec.el, { style, cls })
    const range = positions.of(spec.el)
    if (!range) continue
    const { from, to } = range
    decos.push(Decoration.node(from, to, { class: cls, style }, { key: `col-${cls}-${style}` }))
    const split = splitOf.get(spec.el)
    if (split)
      decos.push(
        Decoration.widget(from + 1, () => splitFloatEl(split), {
          side: -1,
          key: `col-split-${split}`,
          ignoreSelection: true,
        }),
      )
    if (spec.rowFragments) {
      decos.push(...rowDecorations(view, spec.el, spec.rowFragments))
      const transform = `translate(${round2(spec.dx + (spec.pageDx ?? 0))}px, ${round2(spec.dy)}px)`
      // side -3: right after the table, ahead of the page gaps' pull spacer and gap
      decos.push(
        Decoration.widget(to, () => ghostHostEl(transform), {
          side: -3,
          key: `col-host-${transform}`,
          ignoreSelection: true,
        }),
      )
      splitTables.push({ el: spec.el, fragments: spec.rowFragments })
    }
  }
  const next = DecorationSet.create(view.state.doc, decos)
  const prev = key.getState(view.state)
  if (!prev || !sameCols(prev, next)) {
    if (batch) batch.set(key, next)
    else view.dispatch(view.state.tr.setMeta(key, next).setMeta('addToHistory', false))
  }
  const patch = () => {
    patchNodeViews(view, styleOf, splitOf)
    syncSplitTables(view, splitTables)
  }
  if (batch) batch.then(patch)
  else patch()
}

const ROW_CLASS = 'doc-col-row'

function rowDecorations(
  view: EditorView,
  table: HTMLElement,
  fragments: ColumnRowFragment[],
): Decoration[] {
  const decos: Decoration[] = []
  const rows = outerTableRows(table)
  for (const f of fragments) {
    const style = `--row-dx:${round2(f.dx)}px;--row-dy:${round2(f.dy)}px`
    for (const tr of rows.slice(f.startRow, f.endRow)) {
      try {
        const $inside = view.state.doc.resolve(view.posAtDOM(tr, 0))
        for (let d = $inside.depth; d > 0; d--) {
          if ($inside.node(d).type.name !== 'docTableRow') continue
          decos.push(
            Decoration.node(
              $inside.before(d),
              $inside.after(d),
              { class: ROW_CLASS, style },
              { key: `col-row-${style}` },
            ),
          )
          break
        }
      } catch {
        // unmapped DOM: the row stays with its table
      }
    }
  }
  return decos
}

function ghostHostEl(transform: string): HTMLElement {
  const host = document.createElement('div')
  host.className = 'page-float-host doc-col-host'
  host.contentEditable = 'false'
  host.setAttribute('aria-hidden', 'true')
  host.style.transform = transform
  return host
}

/**
 * DOM-only part of a column-split table, after the decorations landed: page
 * gap rows and their repeated header clones (widgets) move with the rows they
 * open, and the host after the table gets one ghost clone per fragment that
 * draws its borders (Chromium paints collapsed borders with the table, where
 * the rows were, so moved rows leave theirs transparent) plus a header clone
 * where the column repeats the header rows. Observer paused like patchNodeViews.
 */
function syncSplitTables(
  view: EditorView,
  tables: Array<{ el: HTMLElement; fragments: ColumnRowFragment[] }>,
): void {
  const obs = (view as unknown as { domObserver?: { stop(): void; start(): void } }).domObserver
  obs?.stop()
  try {
    const moved = new Set<HTMLElement>()
    const zoom = canvasZoomOf(view.dom.closest<HTMLElement>('.doc-zoom'), 1)
    for (const { el, fragments } of tables) {
      const host = el.nextElementSibling
      if (!(host instanceof HTMLElement) || !host.classList.contains('doc-col-host')) continue
      const rows = outerTableRows(el)
      const shiftOf = new Map<HTMLElement, ColumnRowFragment>()
      for (const f of fragments) {
        const own = rows.slice(f.startRow, f.endRow)
        if (own.length === 0) continue
        for (const tr of [...leadWidgetRows(own[0]), ...own]) shiftOf.set(tr, f)
      }
      for (const [tr, f] of shiftOf) {
        if (!tr.matches('.page-gap, .page-repeat-header')) continue
        tr.classList.add(ROW_CLASS)
        tr.style.setProperty('--row-dx', `${round2(f.dx)}px`)
        tr.style.setProperty('--row-dy', `${round2(f.dy)}px`)
        moved.add(tr)
      }
      const geom = tableGeometry(el, host, shiftOf, zoom)
      const ghosts: HTMLElement[] = []
      fragments.forEach((f, i) => {
        const own = rows.slice(f.startRow, f.endRow)
        if (own.length === 0) return
        const lead = leadWidgetRows(own[0])
        const last = own[own.length - 1]
        const next = fragments[i + 1]
        const cutPx = next?.startRow === f.endRow ? next.carriedPx : 0
        // the edge shared with the row above went transparent with the moved row
        // (a row cut at the turn has no bottom edge in this column, as in the preview)
        const above = rows[f.startRow - 1]
        if (above && lead.length === 0 && !shiftOf.has(above) && f.carriedPx < 0.5)
          ghosts.push(ghostTable(el, geom, 0, 0, above, above, 'edge'))
        const first = lead.find((tr) => !tr.classList.contains('page-gap')) ?? own[0]
        ghosts.push(ghostTable(el, geom, f.dx, f.dy, first, last, 'borders', cutPx))
        const header = f.headerPx > 0 ? headerRows(rows) : []
        if (header.length > 0) {
          const end = geom.rows.get(header[header.length - 1])
          const start = geom.rows.get(own[0])
          if (end && start) {
            const dy = start.top + f.dy - (end.top + end.height)
            ghosts.push(
              ghostTable(el, geom, f.dx, dy, header[0], header[header.length - 1], 'header'),
            )
          }
        }
      })
      host.replaceChildren(...ghosts)
    }
    for (const tr of Array.from(view.dom.querySelectorAll<HTMLElement>(`tr.${ROW_CLASS}`))) {
      if (moved.has(tr) || !tr.matches('.page-gap, .page-repeat-header')) continue
      tr.classList.remove(ROW_CLASS)
      tr.style.removeProperty('--row-dx')
      tr.style.removeProperty('--row-dy')
    }
  } finally {
    obs?.start()
  }
}

/** page gap row and repeated header clones a page turn put right above a row */
function leadWidgetRows(tr: HTMLElement): HTMLElement[] {
  const lead: HTMLElement[] = []
  let prev = tr.previousElementSibling
  while (prev instanceof HTMLElement && prev.matches('.page-gap, .page-repeat-header')) {
    lead.unshift(prev)
    prev = prev.previousElementSibling
  }
  return lead
}

function headerRows(rows: HTMLElement[]): HTMLElement[] {
  const end = rows.findIndex((tr) => tr.getAttribute('data-repeat-header') !== '1')
  return rows.slice(0, end < 0 ? rows.length : end)
}

interface TableGeometry {
  /** host origin in the table's box (px) */
  hostTop: number
  hostLeft: number
  height: number
  /** every tbody row's untranslated top and height in the table's box (px) */
  rows: Map<HTMLElement, { top: number; height: number }>
}

/** Fractional layout from rects (offsetTop rounds): the host and the table
 *  share the column translate, and each moved row's own translate is known. */
function tableGeometry(
  table: HTMLElement,
  host: HTMLElement,
  shiftOf: Map<HTMLElement, ColumnRowFragment>,
  zoom: number,
): TableGeometry {
  const t = table.getBoundingClientRect()
  const h = host.getBoundingClientRect()
  const rows = new Map<HTMLElement, { top: number; height: number }>()
  for (const tr of Array.from(table.querySelectorAll<HTMLElement>(':scope > tbody > tr'))) {
    const r = tr.getBoundingClientRect()
    rows.set(tr, {
      top: (r.top - t.top) / zoom - (shiftOf.get(tr)?.dy ?? 0),
      height: r.height / zoom,
    })
  }
  return {
    hostTop: (h.top - t.top) / zoom,
    hostLeft: (h.left - t.left) / zoom,
    height: t.height / zoom,
    rows,
  }
}

/** half the thickest collapsed border on one side of a row (px) */
function halfBorderPx(tr: HTMLElement, side: 'Top' | 'Bottom'): number {
  let w = 0
  for (const cell of Array.from(tr.children))
    w = Math.max(w, parseFloat(getComputedStyle(cell)[`border${side}Width`]) || 0)
  return w / 2
}

/**
 * Clone of the table placed like its moved rows, clipped to rows from..to,
 * less the last row's part cut into the next column (cutPx). A header ghost
 * keeps only the header rows; a border ghost keeps every row at its height
 * with empty cells, so no page gap, picture or paragraph is copied.
 */
function ghostTable(
  table: HTMLElement,
  geom: TableGeometry,
  dx: number,
  dy: number,
  from: HTMLElement,
  to: HTMLElement,
  kind: 'borders' | 'header' | 'edge',
  cutPx = 0,
): HTMLElement {
  const ghost = table.cloneNode(true) as HTMLElement
  ghost.removeAttribute('data-idx')
  ghost.classList.remove('doc-col-block')
  ghost.classList.add('doc-col-ghost', `doc-col-ghost-${kind}`)
  for (const el of Array.from(ghost.querySelectorAll('[data-base-rowspan]')))
    el.removeAttribute('data-base-rowspan')
  const real = Array.from(table.querySelectorAll<HTMLElement>(':scope > tbody > tr'))
  const copies = Array.from(ghost.querySelectorAll<HTMLElement>(':scope > tbody > tr'))
  const lastKept = real.indexOf(to)
  copies.forEach((copy, i) => {
    if (kind === 'header') {
      if (i > lastKept) copy.remove()
      else copy.classList.remove(ROW_CLASS)
      return
    }
    copy.removeAttribute('class')
    copy.style.cssText = `height:${geom.rows.get(real[i])?.height ?? 0}px`
    for (const cell of Array.from(copy.children)) cell.replaceChildren()
  })
  const fromBox = geom.rows.get(from) ?? { top: 0, height: 0 }
  const toBox = geom.rows.get(to) ?? fromBox
  const half = halfBorderPx(to, 'Bottom')
  const bottom = toBox.top + toBox.height + (cutPx > 0 ? -cutPx : half)
  const top = kind === 'edge' ? bottom - 2 * half : fromBox.top - halfBorderPx(from, 'Top')
  ghost.style.left = `${round2(dx - geom.hostLeft)}px`
  ghost.style.top = `${round2(dy - geom.hostTop)}px`
  // a header ghost ends with its header rows; its outer bottom border spills below
  const below = kind === 'header' ? -half : geom.height - bottom
  ghost.style.clipPath = `inset(${round2(top)}px -100vw ${round2(below)}px -100vw)`
  return ghost
}

/**
 * custom NodeViews (protected blocks etc.) don't apply node decorations — patch
 * their DOM directly, observer paused so PM never re-parses the mutation (same
 * technique as the phantom-rowspan sync). Re-applied by every remeasure pass.
 */
function patchNodeViews(
  view: EditorView,
  styleOf: Map<HTMLElement, { style: string; cls: string }>,
  splitOf: Map<HTMLElement, string>,
): void {
  const obs = (view as unknown as { domObserver?: { stop(): void; start(): void } }).domObserver
  obs?.stop()
  try {
    for (const el of Array.from(
      view.dom.querySelectorAll<HTMLElement>('[data-col-patch]'),
    ) as HTMLElement[]) {
      if (styleOf.has(el)) continue
      el.removeAttribute('data-col-patch')
      el.classList.remove(...PATCH_CLASSES)
      for (const p of PATCH_PROPS) el.style.removeProperty(p)
    }
    // split floats the widget decoration could not place (atomic node views)
    for (const el of Array.from(
      view.dom.querySelectorAll<HTMLElement>(`[data-col-split-patch] > .${SPLIT_CLASS}`),
    ) as HTMLElement[]) {
      const host = el.parentElement as HTMLElement
      if (splitOf.get(host) === el.style.cssText) continue
      el.remove()
      host.removeAttribute('data-col-split-patch')
    }
    for (const [el, style] of splitOf) {
      if (el.querySelector(`:scope > .${SPLIT_CLASS}`)) continue
      el.setAttribute('data-col-split-patch', '1')
      el.insertBefore(splitFloatEl(style), el.firstChild)
    }
    for (const [el, { style, cls }] of styleOf) {
      const classes = cls.split(' ')
      if (classes.every((c) => el.classList.contains(c)) && !el.hasAttribute('data-col-patch'))
        continue
      el.setAttribute('data-col-patch', '1')
      el.classList.add(...classes)
      const parts = style.split(';')
      for (const p of PATCH_PROPS) el.style.removeProperty(p)
      for (const part of parts) {
        const i = part.indexOf(':')
        if (i > 0) el.style.setProperty(part.slice(0, i), part.slice(i + 1))
      }
    }
  } finally {
    obs?.start()
  }
}

const round2 = (n: number) => Math.round(n * 100) / 100
const round4 = (n: number) => Math.round(n * 10000) / 10000

function sameCols(a: DecorationSet, b: DecorationSet): boolean {
  const fa = a.find()
  const fb = b.find()
  if (fa.length !== fb.length) return false
  const sig = (d: Decoration) => `${d.from}-${d.to}-${(d.spec as { key?: string }).key ?? ''}`
  const sa = fa.map(sig).sort()
  const sb = fb.map(sig).sort()
  return sa.every((s, i) => s === sb[i])
}
