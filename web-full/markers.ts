// Which document nodes the host's slot markers point at, as ProseMirror node decorations. Pure: no
// view, so the matching rules are tested on a bare document (markersPlugin.ts applies them).
import type { Node as PmNode } from '@tiptap/pm/model'
import { Decoration } from '@tiptap/pm/view'
import type { BlockRef, MarkerTone, SlotMarker } from './embed-protocol'

export const MARKER_CLASS = 'vound-slot'
export const CURRENT_CLASS = 'vound-slot-current'
/** the block under the pointer in review mode (fieldChip.ts) */
export const HOT_CLASS = 'vound-slot-hot'

const cellTypes = ['docTableCell', 'docTableHeader']

/** the tone a node shows when several slots share it: the one that needs attention most comes first */
const tonePriority: MarkerTone[] = ['rejected', 'gap', 'refused', 'conditional', 'proposed', 'accepted']
const attentionRank = (marker: SlotMarker) => tonePriority.indexOf(marker.tone)

type Located = { pos: number; node: PmNode }

/** every node carrying one of the wanted source paragraph indices, by index */
function nodesByIndex(doc: PmNode, wanted: Set<number>) {
  const found = new Map<number, Located[]>()
  doc.descendants((node, pos) => {
    const index = node.attrs.docxPara
    if (typeof index === 'number' && wanted.has(index)) found.set(index, [...(found.get(index) ?? []), { pos, node }])
    // a paragraph's children are inline; nothing below it carries docxPara
    return !node.isTextblock
  })
  return found
}

/** whether any node carries a source paragraph index: a parsed document does, the editor's initial
 *  content and a new blank document do not */
export function hasSourceIndices(doc: PmNode) {
  let found = false
  doc.descendants((node) => {
    if (typeof node.attrs.docxPara === 'number') found = true
    return !found && !node.isTextblock
  })
  return found
}

function enclosingCell(doc: PmNode, pos: number): Located | null {
  const $pos = doc.resolve(pos)
  for (let depth = $pos.depth; depth > 0; depth--) {
    const node = $pos.node(depth)
    if (cellTypes.includes(node.type.name)) return { pos: $pos.before(depth), node }
  }
  return null
}

/** the block a paragraph index and scope name: the paragraph, or the cell around it for cell scope */
function blockTarget(doc: PmNode, scope: BlockRef['scope'], matches: Located[]) {
  // a split copies docxPara to both halves: an index naming two nodes is ambiguous, so it names no block
  if (matches.length !== 1) return null
  return scope === 'cell' ? enclosingCell(doc, matches[0].pos) : matches[0]
}

/** the node a field write or caret placement targets, found as a marker on it would be; null when the
 *  index names no node or several, or a cell scope finds no enclosing cell */
export function locateBlock(doc: PmNode, block: BlockRef): Located | null {
  const found = nodesByIndex(doc, new Set([block.paragraph]))
  return blockTarget(doc, block.scope, found.get(block.paragraph) ?? [])
}

/** one node decoration for every slot on the node: "Ort: ___ Datum: ___" is one paragraph with two
 *  slots. It shows the tone that needs attention most, lists every id in data-vound-slot (space separated,
 *  CSS `~=` matches one) and in its spec, and is current when any of its ids is */
function nodeDecoration(target: Located, markers: SlotMarker[], currentId: string | null) {
  const slotIds = markers.map((marker) => marker.id)
  const shown = markers.reduce((top, marker) => (attentionRank(marker) < attentionRank(top) ? marker : top))
  const current = currentId !== null && slotIds.includes(currentId) ? ` ${CURRENT_CLASS}` : ''
  return Decoration.node(
    target.pos,
    target.pos + target.node.nodeSize,
    { class: `${MARKER_CLASS} ${MARKER_CLASS}-${shown.tone}${current}`, 'data-vound-slot': slotIds.join(' ') },
    { slotIds, isField: markers.some((marker) => marker.actions.length > 0) },
  )
}

/** node decorations for the markers, one per marked node; `dropped` lists ids whose index matches no node
 *  or more than one, or whose cell scope finds no enclosing docTableCell/docTableHeader */
export function markerDecorations(doc: PmNode, markers: SlotMarker[], currentId: string | null) {
  const dropped: string[] = []
  if (markers.length === 0) return { decorations: [], dropped }
  const found = nodesByIndex(doc, new Set(markers.map((marker) => marker.paragraph)))
  const byNode = new Map<number, { target: Located; markers: SlotMarker[] }>()
  for (const marker of markers) {
    const target = blockTarget(doc, marker.scope, found.get(marker.paragraph) ?? [])
    if (!target) {
      dropped.push(marker.id)
      continue
    }
    byNode.set(target.pos, { target, markers: [...(byNode.get(target.pos)?.markers ?? []), marker] })
  }
  const decorations = [...byNode.values()].map((group) => nodeDecoration(group.target, group.markers, currentId))
  return { decorations, dropped }
}
