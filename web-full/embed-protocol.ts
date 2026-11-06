/**
 * postMessage contract between a host page (the Vound app) and this iframe.
 * URL:             index.html?lang=de                                          UI language (de, en, fr, …); unsupported
 *                                                                              values (bg) and none mean English
 * Host → iframe:   { type: 'vound:open', name: string, bytes: ArrayBuffer, mode?: OpenMode }   open a document;
 *                                                                               'review' opens it in review mode,
 *                                                                               'view' read-only (viewMode.ts)
 *                  { type: 'vound:command', command: 'save' }                 save now (same path as ⌘S)
 *                  { type: 'vound:command', command: 'fillWithAi' }           answer with vound:fillWithAi for the caret
 *                  { type: 'vound:markers', markers: SlotMarker[] }            replace the slot markers (all of them)
 *                  { type: 'vound:scrollTo', id: string }                      scroll a marker into view and make it current
 *                  { type: 'vound:focus', id: string | null }                  make a marker current without scrolling, or none
 *                  { type: 'vound:setFieldValue', requestId, block, previous, value, replaces? }
 *                                                                              replace a field's value in one block
 *                  { type: 'vound:setFieldValue', requestId, block, value, replaces: 'gap', gap }
 *                                                                              replace what lies between gap.before
 *                                                                              and gap.after in one block
 *                  { type: 'vound:placeCaret', block: BlockRef }               caret at the end of the block's text
 *                  { type: 'vound:editScope', scope: EditScope }               review mode: what the reviewer may type in (editScope.ts)
 * Iframe → host:   { type: 'vound:ready' }                                     the editor subscribed to open requests
 *                  { type: 'vound:saved', name: string, bytes: ArrayBuffer }   every save (⌘S, File › Save, 'save' command)
 *                  { type: 'vound:dirty', dirty: boolean }                     unsaved changes appeared or were saved; sent on change
 *                  { type: 'vound:error', message: string }                    open or save failed
 *                  { type: 'vound:fieldAction', id, action: FieldAction }      a chip button was pressed (id = slot id); in
 *                                                                              review mode also Mod-Enter (accept) and
 *                                                                              Mod-Shift-Backspace (reject) with the caret
 *                                                                              in a field offering that action (Mod = ⌘
 *                                                                              on the Mac, Ctrl elsewhere)
 *                  { type: 'vound:blockEdited', ids: string[], texts: EditedText[] }
 *                                                                              a user edit changed the text of a marked
 *                                                                              block; texts: each id's block text before
 *                                                                              and after it
 *                  { type: 'vound:caretBlock', ids: string[], leftWith? }      the caret entered a marked block, or left
 *                                                                              one (ids [], leftWith 'Tab' | 'other')
 *                  { type: 'vound:fieldValueResult', requestId, ok, reason?, gap? }
 *                                                                              the answer to vound:setFieldValue; gap
 *                                                                              anchors the written place for the next write
 *                  { type: 'vound:fillWithAi', ids, selected, paragraph }      the reviewer asked AI to fill the caret's place
 *                  { type: 'vound:editRefused' }                               an edit outside the fields was refused (scope 'fields')
 *
 * Both sides check the sender. The editor accepts a message only when event.origin is its own origin
 * and event.source is window.parent (embedMessages.ts); the host should accept one only from the
 * iframe's contentWindow on its own origin. Both post with targetOrigin = window.location.origin, so
 * host and editor have to be served from the same origin. Bytes travel as transferable ArrayBuffers.
 */
export type MarkerTone = 'proposed' | 'gap' | 'conditional' | 'refused' | 'accepted' | 'rejected' | 'otherParty'
/** What a field's chip offers; the host sends the set per field. */
export type FieldAction = 'accept' | 'reject' | 'fill' | 'ask' | 'restore'
/** A highlight on one paragraph of the opened document, named by its source paragraph index: the
 *  paragraph's w:p among all w:p of the original document.xml. `cell` highlights the table cell
 *  around that paragraph instead. `label` names the field on its chip and `actions` are the chip's
 *  buttons; a host that omits them gets `label: id` and `actions: []`. Display only: markers never
 *  reach the saved bytes. */
export type SlotMarker = {
  id: string
  paragraph: number
  scope: 'paragraph' | 'cell'
  tone: MarkerTone
  label: string
  actions: FieldAction[]
  /** 'toggle': the field is a box, so locked to the fields the words on its line stay as they are */
  entry?: 'toggle'
}
/** The block a field write or caret placement targets, named like a marker: the source paragraph
 *  index, and whether the block is that paragraph or the table cell around it. */
export type BlockRef = { paragraph: number; scope: 'paragraph' | 'cell' }
/** 'fields': only the blocks of markers offering actions take edits. 'everything': the whole document does. */
export type EditScope = 'fields' | 'everything'
export type FieldValueRefusal = 'absent' | 'ambiguous' | 'notEditable' | 'noBlock'
/** The visible text of a block right before and right after a field's place, as much as makes the place
 *  unique; '' anchors to the start or the end of the block's text (paragraphs joined by line feeds). */
export type FieldGap = { before: string; after: string }
/** A field's visible block text (its marked blocks joined by line feeds, '' once they are gone) right before
 *  and right after the reviewer's edit; an undo that takes the typing back brings `after` back to the text
 *  from before it. */
export type EditedText = { id: string; before: string; after: string }
/** 'review': the host's review of a filled form (reviewMode.ts). 'view': the document shown, never changed (viewMode.ts). */
export type OpenMode = 'review' | 'view'
export type HostToEditor =
  | { type: 'vound:open'; name: string; bytes: ArrayBuffer; mode?: OpenMode }
  | { type: 'vound:command'; command: 'save' | 'fillWithAi' }
  | { type: 'vound:markers'; markers: SlotMarker[] }
  | { type: 'vound:scrollTo'; id: string }
  | { type: 'vound:focus'; id: string | null }
  | { type: 'vound:setFieldValue'; requestId: string; block: BlockRef; previous: string; value: string; replaces?: 'wholeField' }
  | { type: 'vound:setFieldValue'; requestId: string; block: BlockRef; value: string; replaces: 'gap'; gap: FieldGap }
  | { type: 'vound:placeCaret'; block: BlockRef }
  | { type: 'vound:editScope'; scope: EditScope }
export type EditorToHost =
  | { type: 'vound:ready' }
  | { type: 'vound:saved'; name: string; bytes: ArrayBuffer }
  | { type: 'vound:dirty'; dirty: boolean }
  | { type: 'vound:error'; message: string }
  | { type: 'vound:fieldAction'; id: string; action: FieldAction }
  | { type: 'vound:blockEdited'; ids: string[]; texts: EditedText[] }
  | { type: 'vound:caretBlock'; ids: string[]; leftWith?: 'Tab' | 'other' }
  | { type: 'vound:fillWithAi'; ids: string[]; selected: string; paragraph: string }
  | { type: 'vound:fieldValueResult'; requestId: string; ok: boolean; reason?: FieldValueRefusal; gap?: FieldGap }
  | { type: 'vound:editRefused' }
export const isEmbedded = (): boolean => window.parent !== window
