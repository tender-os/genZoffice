/**
 * postMessage contract between a host page (the Vound app) and this iframe.
 * URL:             index.html?lang=de                                          UI language (de, en, fr, …); unsupported
 *                                                                              values (bg) and none mean English
 * Host → iframe:   { type: 'vound:open', name: string, bytes: ArrayBuffer, mode?: OpenMode }   open a document;
 *                                                                               'review' opens it in review mode,
 *                                                                               'view' read-only (viewMode.ts)
 *                  { type: 'vound:command', command: 'save' }                 save now (same path as ⌘S)
 *                  { type: 'vound:command', command: 'fillWithAi' }           answer with vound:fillWithAi for the caret
 *                  { type: 'vound:command', command: 'undo' | 'redo' }        undo or redo, and focus the document
 *                  { type: 'vound:command', command: 'zoomIn' | 'zoomOut' | 'zoomFitWidth' }
 *                                                                              zoom a step, or to the page width (it then
 *                                                                              follows resizes until the next zoom)
 *                  { type: 'vound:setZoom', zoom: number }                     zoom to this percent, clamped to the range
 *                  { type: 'vound:markers', markers: SlotMarker[] }            replace the slot markers (all of them)
 *                  { type: 'vound:scrollTo', id: string }                      scroll a marker into view and make it current
 *                  { type: 'vound:focus', id: string | null }                  make a marker current without scrolling, or none
 *                  { type: 'vound:setFieldValue', requestId, block, previous, value, replaces? }
 *                                                                              replace a field's value in one block
 *                  { type: 'vound:setFieldValue', requestId, block, value, replaces: 'gap', gap }
 *                                                                              replace what lies between gap.before
 *                                                                              and gap.after in one block
 *                  { type: 'vound:setFieldStyle', requestId, block, previous, replaces?, formText?, style }
 *                  { type: 'vound:setFieldStyle', requestId, block, replaces: 'gap', gap, style }
 *                                                                              restyle a field's text: the selection inside
 *                                                                              it, else all of it (fieldStyle.ts)
 *                  { type: 'vound:placeCaret', block: BlockRef }               caret at the end of the block's text
 *                  { type: 'vound:settleField', block: BlockRef }              the block's legacy text fields are accepted or
 *                                                                              left blank: they lose the author's highlight
 *                                                                              (fieldSettle.ts), an undoable host write
 *                  { type: 'vound:reopenField', block: BlockRef }              the block's settled fields are open again: the
 *                                                                              author's highlight comes back
 *                  { type: 'vound:editScope', scope: EditScope }               review mode: what the reviewer may type in (editScope.ts)
 *                  { type: 'vound:reveal', requestId, text, pageHint? }        scroll to a cited passage and mark it (reveal.ts);
 *                                                                              pageHint (1-based) picks among several matches
 * Iframe → host:   { type: 'vound:ready' }                                     the editor subscribed to open requests
 *                  { type: 'vound:saved', name: string, bytes: ArrayBuffer }   every save (⌘S, File › Save, 'save' command)
 *                  { type: 'vound:dirty', dirty: boolean }                     unsaved changes appeared or were saved; sent on change
 *                  { type: 'vound:error', message: string }                    open or save failed
 *                  { type: 'vound:fieldAction', id, action: FieldAction, anchor? }
 *                                                                              a chip button was pressed (id = slot id); in
 *                                                                              review mode also Mod-Enter (accept) and
 *                                                                              Mod-Shift-Backspace (reject) with the caret
 *                                                                              in a field offering that action (Mod = ⌘
 *                                                                              on the Mac, Ctrl elsewhere); anchor: where
 *                                                                              the Format button is (action 'format')
 *                  { type: 'vound:blockEdited', ids: string[], texts: EditedText[] }
 *                                                                              a user edit changed the text of a marked
 *                                                                              block; texts: each id's block text before
 *                                                                              and after it
 *                  { type: 'vound:caretBlock', ids: string[], leftWith? }      the caret entered a marked block, or left
 *                                                                              one (ids [], leftWith 'Tab' | 'other')
 *                  { type: 'vound:fieldValueResult', requestId, ok, reason?, gap? }
 *                                                                              the answer to vound:setFieldValue; gap
 *                                                                              anchors the written place for the next write
 *                  { type: 'vound:fieldStyleResult', requestId, ok, reason?, style? }
 *                                                                              the answer to vound:setFieldStyle; style: what
 *                                                                              the restyled text shows now
 *                  { type: 'vound:fillWithAi', ids, selected, paragraph }      the reviewer asked AI to fill the caret's place
 *                  { type: 'vound:editRefused', where }                         an edit was refused; where: 'boxLine' (the words
 *                                                                              on a box's line) or 'outsideFields'
 *                  { type: 'vound:history', canUndo: boolean, canRedo: boolean }
 *                                                                              what there is to undo or redo: once the
 *                                                                              editor is ready, then on each change
 *                  { type: 'vound:zoom', zoom: number }                        the zoom in percent, on each change
 *                  { type: 'vound:revealResult', requestId, found: boolean }   the answer to vound:reveal, once the opened
 *                                                                              document is in the editor
 *                  { type: 'vound:opened' }                                    the document of the last vound:open is in the
 *                                                                              editor (once per open)
 *
 * Both sides check the sender. The editor accepts a message only when event.origin is its own origin
 * and event.source is window.parent (embedMessages.ts); the host should accept one only from the
 * iframe's contentWindow on its own origin. Both post with targetOrigin = window.location.origin, so
 * host and editor have to be served from the same origin. Bytes travel as transferable ArrayBuffers.
 */
export type MarkerTone = 'proposed' | 'gap' | 'conditional' | 'refused' | 'accepted' | 'rejected' | 'otherParty'
/** What a field's chip offers; the host sends the set per field. */
export const FIELD_ACTIONS = ['accept', 'reject', 'fill', 'ask', 'restore', 'blank', 'format'] as const
export type FieldAction = (typeof FIELD_ACTIONS)[number]
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
  /** the buttons' titles in the host's language; a button without one keeps the editor's own, German, title */
  actionTitles?: Partial<Record<FieldAction, string>>
  /** 'toggle': the field is a box, so locked to the fields the words on its line stay as they are. 'ownText':
   *  text the reviewer added or changed outside a field, theirs to type in; it offers no chip */
  entry?: MarkerEntry
}
/** The block a field write or caret placement targets, named like a marker: the source paragraph
 *  index, and whether the block is that paragraph or the table cell around it. */
export type BlockRef = { paragraph: number; scope: 'paragraph' | 'cell' }
/** 'fields': only the blocks of markers offering actions take edits. 'everything': the whole document does. */
export type EditScope = 'fields' | 'everything'
export type FieldValueRefusal = 'absent' | 'ambiguous' | 'notEditable' | 'noBlock'
/** Where a chip button sits in the editor's viewport, in CSS pixels: the host opens its Format menu beside it. */
export type FieldAnchor = { left: number; top: number; width: number; height: number }
/** A field's text style: the Latin face (w:rFonts w:ascii, w:hAnsi) and the size in points (w:sz, w:szCs). */
export const FIELD_ALIGNMENTS = ['left', 'center', 'right'] as const
/** a field paragraph's alignment (w:jc) */
export type FieldAlignment = (typeof FIELD_ALIGNMENTS)[number]
export const FIELD_VERTICAL_ALIGNMENTS = ['top', 'center', 'bottom'] as const
/** the vertical alignment of the table cell around a field (w:vAlign) */
export type FieldVerticalAlignment = (typeof FIELD_VERTICAL_ALIGNMENTS)[number]
/** A restyle: `fontSize: null` takes the field's own size away, so it shows the document's again; the vertical
 *  alignment applies to a field in a table cell only. */
export type FieldTextStyle = {
  fontFamily?: string
  fontSize?: number | null
  alignment?: FieldAlignment
  verticalAlignment?: FieldVerticalAlignment
}
/** What a field's text shows: face and size, `sizing` whether the size is the field's own or the document's, the
 *  paragraph's alignment, and the cell's vertical alignment when the field is in a table cell. A property shown
 *  in more than one way is left out. */
export type FieldShownStyle = {
  fontFamily?: string
  fontSize?: number
  sizing?: 'own' | 'document'
  alignment?: FieldAlignment
  verticalAlignment?: FieldVerticalAlignment
}
/** The visible text of a block right before and right after a field's place, as much as makes the place
 *  unique; '' anchors to the start or the end of the block's text (paragraphs joined by line feeds). */
export type FieldGap = { before: string; after: string }
/** A field's visible block text (its marked blocks joined by line feeds, '' once they are gone) right before
 *  and right after the reviewer's edit; an undo that takes the typing back brings `after` back to the text
 *  from before it. */
export type EditedText = { id: string; before: string; after: string }
/** 'review': the host's review of a filled form (reviewMode.ts). 'view': the document shown, never changed (viewMode.ts). */
export type OpenMode = 'review' | 'view'
export const MARKER_ENTRIES = ['toggle', 'ownText'] as const
export type MarkerEntry = (typeof MARKER_ENTRIES)[number]
/** why an edit was refused: it changed the words on a box's line, or text outside every field */
export const REFUSAL_PLACES = ['boxLine', 'outsideFields'] as const
export type RefusalPlace = (typeof REFUSAL_PLACES)[number]
export const HOST_COMMANDS = ['save', 'fillWithAi', 'undo', 'redo', 'zoomIn', 'zoomOut', 'zoomFitWidth'] as const
/** what the host may ask the editor to do (vound:command) */
export type HostCommand = (typeof HOST_COMMANDS)[number]
/** the commands the app's own menu runs (hostCommands.ts) */
export type MenuHostCommand = Exclude<HostCommand, 'fillWithAi'>
export type HostToEditor =
  | { type: 'vound:open'; name: string; bytes: ArrayBuffer; mode?: OpenMode }
  | { type: 'vound:command'; command: HostCommand }
  | { type: 'vound:setZoom'; zoom: number }
  | { type: 'vound:markers'; markers: SlotMarker[] }
  | { type: 'vound:scrollTo'; id: string }
  | { type: 'vound:focus'; id: string | null }
  | { type: 'vound:setFieldValue'; requestId: string; block: BlockRef; previous: string; value: string; replaces?: 'wholeField' }
  | { type: 'vound:setFieldValue'; requestId: string; block: BlockRef; value: string; replaces: 'gap'; gap: FieldGap }
  | { type: 'vound:setFieldStyle'; requestId: string; block: BlockRef; previous: string; replaces?: 'wholeField'; formText?: number; style: FieldTextStyle }
  | { type: 'vound:setFieldStyle'; requestId: string; block: BlockRef; replaces: 'gap'; gap: FieldGap; style: FieldTextStyle }
  | { type: 'vound:placeCaret'; block: BlockRef }
  | { type: 'vound:settleField'; block: BlockRef }
  | { type: 'vound:reopenField'; block: BlockRef }
  | { type: 'vound:editScope'; scope: EditScope }
  | { type: 'vound:reveal'; requestId: string; text: string; pageHint?: number }
export type EditorToHost =
  | { type: 'vound:ready' }
  | { type: 'vound:saved'; name: string; bytes: ArrayBuffer }
  | { type: 'vound:dirty'; dirty: boolean }
  | { type: 'vound:error'; message: string }
  | { type: 'vound:fieldAction'; id: string; action: FieldAction; anchor?: FieldAnchor }
  | { type: 'vound:blockEdited'; ids: string[]; texts: EditedText[] }
  | { type: 'vound:caretBlock'; ids: string[]; leftWith?: 'Tab' | 'other' }
  | { type: 'vound:fillWithAi'; ids: string[]; selected: string; paragraph: string }
  | { type: 'vound:fieldValueResult'; requestId: string; ok: boolean; reason?: FieldValueRefusal; gap?: FieldGap }
  | { type: 'vound:fieldStyleResult'; requestId: string; ok: boolean; reason?: FieldValueRefusal; style?: FieldShownStyle }
  | { type: 'vound:editRefused'; where: RefusalPlace }
  | { type: 'vound:history'; canUndo: boolean; canRedo: boolean }
  | { type: 'vound:zoom'; zoom: number }
  | { type: 'vound:revealResult'; requestId: string; found: boolean }
  | { type: 'vound:opened' }
export const isEmbedded = (): boolean => window.parent !== window
