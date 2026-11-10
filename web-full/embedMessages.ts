// The editor's side of the embed contract (embed-protocol.ts): which incoming messages it trusts and
// how it reads them. Pure, so the rules are tested without a browser.
import { FIELD_ACTIONS, HOST_COMMANDS, MARKER_ENTRIES, type BlockRef, type FieldAction, type FieldGap, type HostToEditor, type MarkerTone, type OpenMode, type SlotMarker } from './embed-protocol'

/** Only the embedding page, and only on the editor's own origin, may drive the editor. */
export const acceptsHostMessage = (
  event: { origin: string; source: unknown },
  self: { origin: string; parent: Window },
) => event.origin === self.origin && event.source === self.parent

const isRecord = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null

const markerTones: MarkerTone[] = ['proposed', 'gap', 'conditional', 'refused', 'accepted', 'rejected', 'otherParty']
const openModes: OpenMode[] = ['review', 'view']

const isSourceIndex = (value: unknown): value is number => typeof value === 'number' && Number.isInteger(value) && value >= 0
const isScope = (value: unknown): value is BlockRef['scope'] => value === 'paragraph' || value === 'cell'

/** A chip's buttons: absent means none; an action this editor does not know is left out, so a newer host
 *  keeps the buttons it has. */
function readActions(value: unknown): FieldAction[] | null {
  if (value === undefined) return []
  if (!Array.isArray(value)) return null
  return value.flatMap((entry) => FIELD_ACTIONS.filter((known) => known === entry))
}

/** The buttons' titles the host gives as text, for actions this editor knows; undefined without any. */
function readActionTitles(value: unknown): SlotMarker['actionTitles'] {
  if (!isRecord(value)) return undefined
  const titles: Partial<Record<FieldAction, string>> = {}
  for (const action of FIELD_ACTIONS) {
    const title = value[action]
    if (typeof title === 'string') titles[action] = title
  }
  return Object.keys(titles).length > 0 ? titles : undefined
}

/** One marker with exactly the known fields, or null when a field is missing or out of range. A host
 *  that predates labels and chip actions gets the id as label and no actions. */
function readMarker(value: unknown): SlotMarker | null {
  if (!isRecord(value)) return null
  const { id, paragraph, scope, tone, label } = value
  if (typeof id !== 'string') return null
  if (!isSourceIndex(paragraph) || !isScope(scope)) return null
  const knownTone = markerTones.find((known) => known === tone)
  if (!knownTone) return null
  if (label !== undefined && typeof label !== 'string') return null
  const actions = readActions(value.actions)
  if (!actions) return null
  const actionTitles = readActionTitles(value.actionTitles)
  const marker = { id, paragraph, scope, tone: knownTone, label: label ?? id, actions, ...(actionTitles ? { actionTitles } : {}) }
  const entry = MARKER_ENTRIES.find((known) => known === value.entry)
  return entry ? { ...marker, entry } : marker
}

/** The block a field write or caret placement names, with exactly the known fields. */
function readBlock(value: unknown): BlockRef | null {
  if (!isRecord(value)) return null
  const { paragraph, scope } = value
  return isSourceIndex(paragraph) && isScope(scope) ? { paragraph, scope } : null
}

/** All markers, or null when the list is not an array or any one entry is malformed. */
function readMarkers(value: unknown): SlotMarker[] | null {
  if (!Array.isArray(value)) return null
  const markers = value.map(readMarker)
  return markers.every((marker): marker is SlotMarker => marker !== null) ? markers : null
}

function readGap(value: unknown): FieldGap | null {
  if (!isRecord(value) || typeof value.before !== 'string' || typeof value.after !== 'string') return null
  return { before: value.before, after: value.after }
}

/** The host message in `data`, or null when it is not one the editor accepts. */
export function readHostMessage(data: unknown): HostToEditor | null {
  if (!isRecord(data)) return null
  if (data.type === 'vound:open' && typeof data.name === 'string' && data.bytes instanceof ArrayBuffer) {
    // an unknown mode opens the document as without one
    const open = { type: 'vound:open', name: data.name, bytes: data.bytes } as const
    const mode = openModes.find((known) => known === data.mode)
    return mode ? { ...open, mode } : open
  }
  const command = HOST_COMMANDS.find((known) => known === data.command)
  if (data.type === 'vound:command' && command) return { type: 'vound:command', command }
  if (data.type === 'vound:setZoom' && typeof data.zoom === 'number' && Number.isFinite(data.zoom)) {
    return { type: 'vound:setZoom', zoom: data.zoom }
  }
  if (data.type === 'vound:markers') {
    const markers = readMarkers(data.markers)
    return markers ? { type: 'vound:markers', markers } : null
  }
  if (data.type === 'vound:scrollTo' && typeof data.id === 'string') return { type: 'vound:scrollTo', id: data.id }
  if (data.type === 'vound:focus' && (typeof data.id === 'string' || data.id === null)) return { type: 'vound:focus', id: data.id }
  if (data.type === 'vound:setFieldValue' && data.replaces === 'gap') {
    const { requestId, value } = data
    const block = readBlock(data.block)
    const gap = readGap(data.gap)
    if (!block || !gap || typeof requestId !== 'string' || typeof value !== 'string') return null
    return { type: 'vound:setFieldValue', requestId, block, value, replaces: 'gap', gap }
  }
  if (data.type === 'vound:setFieldValue') {
    const { requestId, previous, value } = data
    const block = readBlock(data.block)
    if (!block || typeof requestId !== 'string' || typeof previous !== 'string' || typeof value !== 'string') return null
    const write = { type: 'vound:setFieldValue', requestId, block, previous, value } as const
    return data.replaces === 'wholeField' ? { ...write, replaces: 'wholeField' } : write
  }
  if (data.type === 'vound:placeCaret') {
    const block = readBlock(data.block)
    return block ? { type: 'vound:placeCaret', block } : null
  }
  if (data.type === 'vound:editScope' && (data.scope === 'fields' || data.scope === 'everything')) {
    return { type: 'vound:editScope', scope: data.scope }
  }
  if (data.type === 'vound:reveal') return readReveal(data)
  return null
}

/** A reveal needs its request id and some text; a page hint is a positive whole page number or left out. */
function readReveal(data: Record<string, unknown>): HostToEditor | null {
  const { requestId, text, pageHint } = data
  if (typeof requestId !== 'string' || typeof text !== 'string' || !text.trim()) return null
  const reveal = { type: 'vound:reveal', requestId, text } as const
  if (pageHint === undefined) return reveal
  return isSourceIndex(pageHint) && pageHint > 0 ? { ...reveal, pageHint } : null
}
