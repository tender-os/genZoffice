// The editor's side of the embed contract (embed-protocol.ts): which incoming messages it trusts and
// how it reads them. Pure, so the rules are tested without a browser.
import type { BlockRef, FieldAction, HostToEditor, MarkerTone, SlotMarker } from './embed-protocol'

/** Only the embedding page, and only on the editor's own origin, may drive the editor. */
export const acceptsHostMessage = (
  event: { origin: string; source: unknown },
  self: { origin: string; parent: Window },
) => event.origin === self.origin && event.source === self.parent

const isRecord = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null

const markerTones: MarkerTone[] = ['proposed', 'gap', 'conditional', 'refused', 'accepted', 'rejected', 'otherParty']
const fieldActions: FieldAction[] = ['accept', 'reject', 'fill', 'ask', 'restore']

const isSourceIndex = (value: unknown): value is number => typeof value === 'number' && Number.isInteger(value) && value >= 0
const isScope = (value: unknown): value is BlockRef['scope'] => value === 'paragraph' || value === 'cell'

/** A chip's buttons: absent means none; any unknown entry rejects the list. */
function readActions(value: unknown): FieldAction[] | null {
  if (value === undefined) return []
  if (!Array.isArray(value)) return null
  const actions = value.map((entry) => fieldActions.find((known) => known === entry))
  return actions.every((action): action is FieldAction => action !== undefined) ? actions : null
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
  const marker = { id, paragraph, scope, tone: knownTone, label: label ?? id, actions }
  return value.entry === 'toggle' ? { ...marker, entry: 'toggle' } : marker
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

/** The host message in `data`, or null when it is not one the editor accepts. */
export function readHostMessage(data: unknown): HostToEditor | null {
  if (!isRecord(data)) return null
  if (data.type === 'vound:open' && typeof data.name === 'string' && data.bytes instanceof ArrayBuffer) {
    // any mode but 'review' opens the document as without one
    const open = { type: 'vound:open', name: data.name, bytes: data.bytes } as const
    return data.mode === 'review' ? { ...open, mode: 'review' } : open
  }
  if (data.type === 'vound:command' && (data.command === 'save' || data.command === 'fillWithAi')) {
    return { type: 'vound:command', command: data.command }
  }
  if (data.type === 'vound:markers') {
    const markers = readMarkers(data.markers)
    return markers ? { type: 'vound:markers', markers } : null
  }
  if (data.type === 'vound:scrollTo' && typeof data.id === 'string') return { type: 'vound:scrollTo', id: data.id }
  if (data.type === 'vound:focus' && (typeof data.id === 'string' || data.id === null)) return { type: 'vound:focus', id: data.id }
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
  return null
}
