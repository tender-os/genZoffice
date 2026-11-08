// Which document nodes the host's slot markers point at, as ProseMirror node decorations. Pure: no
// view, so the matching rules are tested on a bare document (markersPlugin.ts applies them).
import type { Node as PmNode } from '@tiptap/pm/model'
import { Decoration } from '@tiptap/pm/view'
import type { BlockRef, MarkerTone, SlotMarker } from './embed-protocol'

export const MARKER_CLASS = 'vound-slot'
export const CURRENT_CLASS = 'vound-slot-current'
/** the block under the pointer in review mode (fieldChip.ts) */
export const HOT_CLASS = 'vound-slot-hot'
/** a block the reviewer can still type in while the document is locked to the fields (host.css) */
export const FIELD_CLASS = 'vound-slot-field'
const TOGGLE_CLASS = 'vound-slot-toggle'
/** another field's block sits right above or below this one (host.css keeps their tints apart) */
const OTHER_ABOVE_CLASS = 'vound-slot-other-above'
const OTHER_BELOW_CLASS = 'vound-slot-other-below'

const cellTypes = ['docTableCell', 'docTableHeader']

/** the tone a node shows when several slots share it: the one that needs attention most comes first */
const tonePriority: MarkerTone[] = ['rejected', 'gap', 'refused', 'conditional', 'proposed', 'accepted', 'otherParty']
const attentionRank = (marker: SlotMarker) => tonePriority.indexOf(marker.tone)

type Located = { pos: number; node: PmNode }
type Group = { target: Located; markers: SlotMarker[] }

/** every node carrying one of the wanted source paragraph indices, by index, and the paragraphs continuing
 *  one after a split (continuesPara, set by the editor's split-source-paragraph.ts) */
function nodesByIndex(doc: PmNode, wanted: Set<number>) {
  const found = new Map<number, Located[]>()
  const continuations = new Map<number, Located[]>()
  const add = (map: Map<number, Located[]>, index: number, located: Located) => map.set(index, [...(map.get(index) ?? []), located])
  doc.descendants((node, pos) => {
    const { docxPara, continuesPara } = node.attrs
    if (typeof docxPara === 'number' && wanted.has(docxPara)) add(found, docxPara, { pos, node })
    if (typeof continuesPara === 'number' && wanted.has(continuesPara)) add(continuations, continuesPara, { pos, node })
    // a paragraph's children are inline; nothing below it carries docxPara
    return !node.isTextblock
  })
  return { found, continuations }
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
  const { found } = nodesByIndex(doc, new Set([block.paragraph]))
  return blockTarget(doc, block.scope, found.get(block.paragraph) ?? [])
}

/** one node decoration for every slot on the node: "Ort: ___ Datum: ___" is one paragraph with two
 *  slots. It shows the tone that needs attention most, lists every id in data-vound-slot (space separated,
 *  CSS `~=` matches one) and in its spec, and is current when any of its ids is */
function nodeDecoration(target: Located, markers: SlotMarker[], currentId: string | null, apart: string) {
  const slotIds = markers.map((marker) => marker.id)
  const shown = markers.reduce((top, marker) => (attentionRank(marker) < attentionRank(top) ? marker : top))
  const current = currentId !== null && slotIds.includes(currentId) ? ` ${CURRENT_CLASS}` : ''
  const isField = markers.some((marker) => marker.actions.length > 0)
  const field = isField ? ` ${FIELD_CLASS}` : ''
  const togglesOnly = isField && markers.every((marker) => marker.entry === 'toggle')
  const toggle = togglesOnly ? ` ${TOGGLE_CLASS}` : ''
  return Decoration.node(
    target.pos,
    target.pos + target.node.nodeSize,
    { class: `${MARKER_CLASS} ${MARKER_CLASS}-${shown.tone}${field}${toggle}${current}${apart}`, 'data-vound-slot': slotIds.join(' ') },
    togglesOnly ? { slotIds, isField, togglesOnly } : { slotIds, isField },
  )
}

const idsOf = (group: Group) => group.markers.map((marker) => marker.id).join(' ')
const isCell = (target: Located) => cellTypes.includes(target.node.type.name)

/** the classes of a marked block (not a cell) whose sibling right above or below is marked for other fields */
function apartClasses(group: Group, byStart: Map<number, Group>, byEnd: Map<number, Group>) {
  if (isCell(group.target)) return ''
  const otherField = (neighbour: Group | undefined) => neighbour !== undefined && !isCell(neighbour.target) && idsOf(neighbour) !== idsOf(group)
  const above = otherField(byEnd.get(group.target.pos)) ? ` ${OTHER_ABOVE_CLASS}` : ''
  const below = otherField(byStart.get(group.target.pos + group.target.node.nodeSize)) ? ` ${OTHER_BELOW_CLASS}` : ''
  return `${above}${below}`
}

/** node decorations for the markers, one per marked node; `dropped` lists ids whose index matches no node
 *  or more than one, or whose cell scope finds no enclosing docTableCell/docTableHeader */
export function markerDecorations(doc: PmNode, markers: SlotMarker[], currentId: string | null) {
  const dropped: string[] = []
  if (markers.length === 0) return { decorations: [], dropped }
  const { found, continuations } = nodesByIndex(doc, new Set(markers.map((marker) => marker.paragraph)))
  const byNode = new Map<number, Group>()
  const mark = (target: Located, marker: SlotMarker) =>
    byNode.set(target.pos, { target, markers: [...(byNode.get(target.pos)?.markers ?? []), marker] })
  for (const marker of markers) {
    const target = blockTarget(doc, marker.scope, found.get(marker.paragraph) ?? [])
    if (!target) {
      dropped.push(marker.id)
      continue
    }
    mark(target, marker)
    // a field paragraph split by Enter: the field spans its halves
    if (marker.scope === 'paragraph') for (const half of continuations.get(marker.paragraph) ?? []) mark(half, marker)
  }
  const groups = [...byNode.values()]
  const byEnd = new Map(groups.map((group) => [group.target.pos + group.target.node.nodeSize, group]))
  const decorations = groups.map((group) =>
    nodeDecoration(group.target, group.markers, currentId, apartClasses(group, byNode, byEnd)),
  )
  return { decorations, dropped }
}
