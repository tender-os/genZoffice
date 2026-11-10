# GenOffice Docs as an embeddable web editor (Vound fork)

`web-full/` hosts the Docs app's complete renderer (`apps/docs/src/renderer/main.tsx`, unchanged) in a
plain web page. `shim.ts` stands in for the Electron bridge (`window.desktop`), `boot.ts` installs it and
loads the app, `host.css` and two transforms in `vite.config.ts` adjust the app for use as a viewer:

- header/footer strips take the document's base font (upstream rule only covered the body);
- header/footer strips are not dimmed;
- break-only spacer paragraphs are one line, not two (`.doc-br-line` inline);
- the host decides editability; a document's own `w:documentProtection` lock is not enforced.

Engine change in `packages/docx-engine/src/text-patch.ts` (`patchFieldParagraphText`, called from
`generate.ts` and `patch.ts`): a text edit inside a legacy form field or a content control patches
the `w:t` of the original paragraph and keeps the field machinery. Test:
`packages/docx-engine/tests/field-paragraph-text-patch.test.ts`.

## Build

```sh
npm ci --ignore-scripts        # repo root, Node 22+
npm run build:web-host         # -> web-full/dist (static, relative asset paths, ~22 MB with fonts)
```

## Embed

Serve `web-full/dist` under any path of the host's own origin and put it in an iframe. The contract
is `embed-protocol.ts`. `index.html?lang=de` picks the UI language (`normalizeLang` of
`@genoffice/i18n`: `de`, `en`, `fr` and the other supported languages; an unsupported value such as `bg`,
or none, means English); the stand-alone page stays English.

| direction | message |
|---|---|
| host → editor | `{ type: 'vound:open', name, bytes: ArrayBuffer, mode?: 'review' \| 'view' }` open a document (transfer the buffer); `mode: 'review'` opens it in review mode, `mode: 'view'` read-only (see View mode), any other mode or none in the normal editor |
| host → editor | `{ type: 'vound:command', command: 'save' }` save now, the same path as ⌘S |
| host → editor | `{ type: 'vound:command', command: 'undo' \| 'redo' }` undo or redo, and focus the document |
| host → editor | `{ type: 'vound:command', command: 'zoomIn' \| 'zoomOut' \| 'zoomFitWidth' }` zoom a step, or to the page width, which then follows resizes |
| host → editor | `{ type: 'vound:setZoom', zoom }` zoom to this percent, clamped to the editor's range |
| host → editor | `{ type: 'vound:markers', markers: SlotMarker[] }` replace all slot markers (`{ id, paragraph, scope: 'paragraph' \| 'cell', tone, label, actions }`; `actions` ⊂ `accept`, `reject`, `fill`, `ask`, `restore`, `blank`, any other action left out; `actionTitles?`: the buttons' titles by action, in the host's language, else the editor's German ones; without them a marker gets `label: id`, `actions: []`; `entry?: 'toggle' \| 'ownText'`: a box's line, or text the reviewer wrote outside a field, theirs to type in with no chip) |
| host → editor | `{ type: 'vound:scrollTo', id }` scroll a marker to the middle of the view and make it current |
| host → editor | `{ type: 'vound:focus', id: string \| null }` make a marker current without scrolling, or none |
| host → editor | `{ type: 'vound:setFieldValue', requestId, block: BlockRef, previous, value }` replace a field's value in one block (`BlockRef = { paragraph, scope: 'paragraph' \| 'cell' }`, named like a marker); answered with `vound:fieldValueResult` (see Field writes) |
| host → editor | `{ type: 'vound:setFieldValue', requestId, block: BlockRef, value, replaces: 'gap', gap: { before, after } }` replace what lies between the gap's anchors in one block (see Field writes) |
| host → editor | `{ type: 'vound:placeCaret', block: BlockRef }` put the caret at the end of the block's text, focus the editor and scroll it into view |
| host → editor | `{ type: 'vound:settleField', block: BlockRef }` the block's legacy text fields were accepted or left blank: they lose the form author's highlight on screen and in the saved file, as one host write the reviewer can undo |
| host → editor | `{ type: 'vound:reveal', requestId, text, pageHint? }` scroll to a cited passage and mark it (see Reveal); answered with `vound:revealResult` |
| editor → host | `{ type: 'vound:ready' }` after the editor subscribed to open requests |
| editor → host | `{ type: 'vound:saved', name, bytes: ArrayBuffer }` on every save (⌘S, File › Save, the save command) |
| editor → host | `{ type: 'vound:dirty', dirty: boolean }` when unsaved changes appear or a save clears them |
| editor → host | `{ type: 'vound:error', message }` open or save failed |
| editor → host | `{ type: 'vound:editRefused', where: 'boxLine' \| 'outsideFields' }` an edit was refused: it changed the words on a box's line, or text outside every field |
| editor → host | `{ type: 'vound:history', canUndo, canRedo }` what there is to undo or redo: once the editor is ready, then on each change |
| editor → host | `{ type: 'vound:zoom', zoom }` the zoom in percent, on each change |
| editor → host | `{ type: 'vound:fieldAction', id, action }` a field chip's button was pressed (`id` is the slot id) |
| editor → host | `{ type: 'vound:blockEdited', ids, texts }` a user edit changed the text of a marked block; `ids` are its slot ids, `texts` each id's block text right before and right after the edit (an undo of typing brings `after` back to the text before it). Review mode only; host writes are not reported |
| editor → host | `{ type: 'vound:caretBlock', ids, leftWith? }` the caret entered a marked block, or left one (`ids: []`, `leftWith: 'Tab' \| 'other'`). Review mode only |
| editor → host | `{ type: 'vound:fieldValueResult', requestId, ok, reason?, gap? }` the answer to `vound:setFieldValue` (`reason`: `absent`, `ambiguous`, `notEditable`, `noBlock`; `gap`: the written place's anchors) |
| editor → host | `{ type: 'vound:revealResult', requestId, found }` the answer to `vound:reveal`, once the opened document is in the editor |
| editor → host | `{ type: 'vound:opened' }` the document of the last `vound:open` is in the editor; once per open |

Host and editor must share one origin. The editor accepts a message only when `event.origin` is its
own origin and `event.source` is `window.parent` (`embedMessages.ts`), reads it structurally and drops
anything else; it posts with `targetOrigin = window.location.origin`. The host should do the same in
reverse: accept only `frame.contentWindow` on its own origin, post with `location.origin`, never `'*'`.

`vound:dirty` comes from the app's own unsaved-changes state (`isDocDirty`, the one the close guard and
autosave read). The `hostReportsDirty` transform in `vite.config.ts` hands it to the shim after every
render, and the shim posts only when the value changed. Opening a document starts clean; a save that
completes reports `dirty: false`.

`embed-example.html` is a working host page. The editor never persists anything itself when embedded;
the host receives the bytes and stores them. Only explicit saves reach the host (⌘S, File › Save, the
`save` command): an auto save (`saveDocx(path, data, auto = true)`) posts nothing and leaves the document
unsaved (`embeddedSave.ts`), because every `vound:saved` becomes a published revision.

## Markers

The host highlights document paragraphs with `vound:markers`. A marker names its paragraph by the
source paragraph index: the position of the `w:p` among all elements with local name `p` in the opened
file's `word/document.xml`, in document order (`mc:Fallback`, comments and CDATA excluded). The engine
stamps that index on every paragraph it models (`docxPara`, DOM `data-src-p`), and `markers.ts` puts
each marker on the one node carrying it, or on the table cell around it for `scope: 'cell'`. Every
marked node gets one node decoration, also when several slots share it ("Ort: ___ Datum: ___" is one
paragraph with two slots): `data-vound-slot` lists all of its ids separated by spaces (select one with
`[data-vound-slot~="<id>"]`), the class carries the tone that needs attention most (rejected, gap,
refused, conditional, proposed, accepted, in that order), and the node is current when any of its ids
is. An index that matches no node or more than one is dropped and logged with
`console.info('[vound] markers dropped', ids)`; so is a cell marker outside a table cell. A split (Enter inside a
paragraph) copies the index to both halves; the editor then keeps it on the first half and marks the second as its
continuation (`continuesPara`, `split-source-paragraph.ts`), and a paragraph-scope marker covers the paragraph and
its continuations, drawn as one field over several blocks. Halves that still share the index side by side, before
the editor marked the continuation, count the same way, so a split never drops the marker. Joining the halves gives the one paragraph back. Paragraphs of nested tables are not editor nodes and cannot be marked.

`markersPlugin.ts` keeps the decorations through edits and walks the document again when the block
count changes (phased loading appends blocks after the first 64, splits, deletions), when an attribute
change dropped a decoration, or while markers wait for their document. The wait ends at the first
build that sees nodes with source indices, so markers that all miss a loaded document do not cost a
walk per edit. A paragraph the editor replaces by a new one keeps its source index (`sourceIndex.ts`): a cut cell
selection (prosemirror-tables fills the cell with a default paragraph) would otherwise leave a paragraph without one,
and an emptied field would lose its marker until the next save. Backspace and Delete in an emptied cell paragraph,
and Delete, Backspace or typing over selected cells, keep the paragraph itself with its properties, as Word does
(`extensions.ts`: `isEmptyAtEdge`, `clearSelectedCells`, `typeOverSelectedCells`). The new paragraph takes the index of the paragraph it
replaced when no other paragraph carries it, in a transaction appended to the edit. Markers sent before the app handed over its editor (`hostEditorHandle` in
`vite.config.ts`) are applied when it does. `vound:scrollTo` and `vound:focus` add
`vound-slot-current`, and `vound:scrollTo` finds a shared node by any of its ids; `host.css` holds the
six tone colours (review mode restyles them, see Field chip).

After a save the app re-parses the saved bytes, so the nodes may carry the saved file's numbering, which
the host's markers do not name. Outside review mode an explicit save therefore seals the markers:
`embeddedSave.ts` clears them before it posts `vound:saved`, and the editor ignores `vound:markers` and
`vound:focus` until the next `vound:open`. In review mode a save does not seal (`markersOnSave`): the
markers stay on the nodes they mark. The app replaces the document with the re-parsed saved bytes only
when the parse differs from the editor's document and nothing changed it since the save; that replace
keeps the markers the host posted after the save (it posts them once the saved file's body skeleton
matches the source's) and builds them on the new document, and drops older ones, which may name
renumbered paragraphs. When no replace follows (a clean save, or an edit after it), the paragraphs keep
their numbering and so do the markers. The host's markers apply after a save as at any other time.

Every `vound:open` resets the markers (`resetMarkers`: none, no current one, not sealed) before
the document loads, so the previous document's indices never land on the new one; markers posted after
the open apply as usual, so the host posts markers after `vound:open`. The first whole-document replace
after a `vound:open` (or the editor's first document) is that open's document, and the markers are built
on it again. Any later whole-document replace clears them: a document that came another way (File › Open
inside the editor, a password-protected file through `loadFile`, a new blank document) does not carry the
host's numbering. So does the rare remount after a refused phased chunk, and a select-all that replaces
everything; the markers come back with the next `vound:markers` or `vound:open`. Markers are view
decorations only; nothing of them is written into a saved file.

Embedded, `boot.ts` sets `html.vound-embedded` and `host.css` hides the Genspark AI dock, the ribbon's
AI commands (the Genspark AI group, Editor, Translate, AI Resolve Comments, AI Revision Summary), the
floating "Ask AI" button on a selection and the AutoSave toggle; the page area takes the full width. The
editor's success toast ("Saved", which only means the bytes went to the host) is hidden too, since the
host reports the real outcome; error toasts still show.

## Review mode

`vound:open` with `mode: 'review'` opens the document in review mode (`reviewMode.ts`): the host shows
its own bar, and the editor edits like Word without the ribbon. `host.css` hides the ribbon, its tab row, the
status bar and the paste-options chip under `html.vound-review`, and pads the top of the page scroll
container (`.editor-scroll`) by `--vound-review-top` (52px) for the host's bar. Typing, deleting, keyboard
undo/redo, clipboard shortcuts, select-all, save, find, zoom and keyboard selection work as usual.

Paste inserts the clipboard's plain text in the formatting of the insertion point (`pasteText.ts`):
`text/plain`, or for an HTML-only clipboard the HTML's text with one line per block. Line breaks and
paragraph ends become line feeds inside the one paragraph, so a paste never adds paragraphs. The plugin
takes the paste DOM event, because the app's own paste handlers are direct editor props that ProseMirror
asks before any plugin, and some of them insert formatted content. A paste that skips the DOM event
(`view.pasteHTML`) meets the app's `handlePaste` first and reaches the plugin's only when that declines,
so review mode closes the app's ways in: the paste-options chip is hidden, and the app's context menu
offers only Paste as Plain Text.

The app's context menu opens in review mode, also on a protected form (the `hostDecidesEditability`
transform lifts its protection gate with the editing lock). The embed leaves entries out (`reviewMenu.ts`,
the `reviewMenu` transform asks the shim's `menuShows` before each one): the AI entries and hiding text
from the AI always, and in review mode formatted Paste, the field commands, New Comment and every command
that changes a form's structure: deleting from or merging, splitting, distributing and resizing a table,
table properties, inserting columns or cells, and list numbering. Inserting rows stays, since forms such as
reference lists need more of them.

Formatting shortcuts work, as in Word: bold, italic, underline, alignment, the font, paragraph and
hyperlink dialogs, line spacing, paragraph styles, font size, super- and subscript, indents, clearing
formatting and change case. Document-tool shortcuts do nothing (`reviewShortcuts.ts`): fields, track
changes, comments, footnotes and endnotes, AI proofread and File › Open. A window capture listener stops them before the editors' keymaps (main
editor, header and footer, text boxes) and App.tsx's window handler, which acts without looking at
`defaultPrevented`. On the Mac, Control+letter is text navigation (⌃E end of line): the app and the
editor do not see it, the browser's own action stays.

A review-mode open that arrives before the app handed over its editor takes effect at the hand-over.
An open without `mode: 'review'` leaves review mode. File › Open (⌘O) is blocked too: its file would
replace the reviewed document.

Review mode reports the reviewer's work in marked blocks (`fieldReporting.ts`). `vound:blockEdited`
follows a transaction of theirs that changes a marked block (a paragraph or cell carrying a marker
decoration), with the ids of every block it touched, before or after the change (Backspace that joins a
marked paragraph into the one above reports it). Not reported: host writes and transactions other plugins
append to them, document opens, changes kept out of the undo history (the app's TOC, pagination and
formatting refreshes), and a paragraph added or removed between two marked blocks. Select-all and typing
replaces the whole document: that is not reported either, and it drops every marker; the host's
skeleton check of the saved bytes catches it. `vound:caretBlock` follows the caret: the ids of the marked
block it entered, or `ids: []` with `leftWith: 'other'` when it left one. Tab neither indents, inserts a
tab nor moves to the next table cell; in a marked block it reports `ids: []` with `leftWith: 'Tab'`, the
host's cue to move to the next field. Shift+Tab does nothing.

With the caret in a marked block, Mod-Enter posts `vound:fieldAction` with `accept` and Mod-Shift-Backspace
with `reject` (`fieldActionKeys.ts`; Mod is ⌘ on the Mac, Ctrl elsewhere), like the chip's buttons. A key
acts only when the field offers that action: the host's current field if it is at the caret and offers it,
else the only field at the caret that does. Otherwise the key keeps its normal meaning (Mod-Enter inserts
a page break). Inside a legacy form field being edited the node view keeps the keys, as it does for typing.

A field passthrough paragraph (a legacy form field outside a table) is typed into inside its node view, out
of ProseMirror's reach: its text reaches the document when the editor commits the field (the caret leaves
the field, Tab, or a save starts), so its `vound:blockEdited` comes then. While the DOM caret is in a field
being edited, the caret counts as in that field's block. Tab there is caught before the node view: it
commits the field, so `vound:blockEdited` comes first, then reports leaving with Tab. A click on such a
field selects it as an object; review mode then refuses typing, Backspace, Delete, paste and cut, which
would replace the block, and the field with it, by plain text.

Word shows an empty legacy form field as five en spaces (U+2002). Typing into such a blank replaces it
wherever the caret sits in it: in a table cell, where the field's result is the paragraph's text, review
mode replaces any paragraph of en spaces only with the first typed text (field paragraphs are not marked
in the editor); in a field passthrough, the editor's
field commit drops the blank (or an empty line's placeholder NBSP) around the typed text, also when the
caret sat in the middle of it. In review mode the blank shows no width (a `vound-blank` decoration,
negative letter spacing in `host.css`), so the caret and the typed text start at the field's edge; the en
spaces stay in the document.

Placeholder text is selected whole when the caret lands in it, as in Word (`placeholderSelect.ts`), so the first
keystroke replaces it and typing never mixes with it: a click anywhere on it (its edges too) selects it before the
click ends, a caret moved by keys selects it from inside only, so it can still be stepped past, and
`vound:placeCaret` selects placeholder text a block ends with. Placeholder text is what the engine marks
`Run.placeholder` (the editor's `ctrlPlaceholder` mark, which typed text never continues): the content of a content
control whose `w:sdtPr` says `w:showingPlcHdr` (inline, around a paragraph, or around a table cell's paragraph), and
the result of a legacy text form field (FORMTEXT) that equals its `w:default` text. Text that only looks like a
hint, a grey character style with no control around it, is the document's own and is never selected.
A form field outside a table (a field passthrough) whose result is its default text (`fieldDisplay.defaultText`)
is entered with that text selected on a click, and by `vound:placeCaret`.

Text typed over placeholder text is the field's own at once (`placeholder-typing.ts` in the editor): it carries
neither the mark nor the placeholder's character style (`isPlaceholderStyleId`: Placeholder Text, Platzhaltertext),
also when a composition (dead keys) writes it into the placeholder's span; undo brings the placeholder back. The
save keeps the content control around it, its properties and tag included, and drops only `w:showingPlcHdr` and
the placeholder style, as Word does: a control around a body paragraph (`patch.ts`) and one around a table cell's
paragraph (`generate.ts`, `regroupCellParagraphs`). A cell edit writes each paragraph back into the place of the
original paragraph it continues (the editor passes its `origin`, from the source paragraph index), so a paragraph
made inside a control by Enter or a paste stays in it, paragraphs removed from a control leave the wrapper around
the rest, and a control whose paragraphs all went keeps one empty paragraph. Such an edit no longer regenerates the
whole table, as long as the cell holds paragraphs only, no pictures, and every paragraph keeps a format the cell
had.

## View mode

`vound:open` with `mode: 'view'` shows the document and nothing that changes it (`viewMode.ts`). `host.css`
hides the ribbon, its tab row and the status bar under `html.vound-view`, and the pages take no pointer input;
the canvas still scrolls. A plugin whose `editable` prop answers false keeps the editor read-only whatever the
app's `setEditable` says, so there is no caret. A window capture listener keeps every key, double-click and
right-click from the app's handlers, so no shortcut, header edit or context menu acts; Mod+S, Mod+P and Mod+O
do not open the browser's dialogs either. A save in view mode posts nothing. An open in another mode leaves view mode.

## Reveal

`{ type: 'vound:reveal', requestId, text, pageHint? }` scrolls to a passage the host cites and marks it
(`reveal.ts`); the editor answers `{ type: 'vound:revealResult', requestId, found }`. The match ignores case,
whitespace, hyphens (also the soft hyphen) and the quote style, and folds ligatures (NFKC), so a quote taken
from the converted PDF finds the Word text across line breaks, hyphenated words and paragraph breaks. Of several
matches, the one on the page nearest `pageHint` (1-based; a page is counted by the `.page-gap` elements laid out
before it) is taken, the first without a hint. The mark is a decoration (`.vound-reveal` in `host.css`): one at a
time, replaced by the next reveal, cleared by the next open, never saved. The editor posts `vound:opened` once
the opened document is in, so a host can time its wait for an answer from then. A reveal that arrives right after
`vound:open` waits until the opened document is in the editor (`revealQueue.ts`); a newer reveal or open answers
the waiting one with `found: false`. A message without a request id or text, or with a page hint that is not a
whole page number, is dropped unanswered. Text of legacy form fields (passthrough blocks) is not searched.

## Field chip

In review mode (`fieldChip.ts`, mounted by `reviewMode.ts`) hovering a marked block shows a chip with the
buttons the host sent for its fields (`SlotMarker.actions`): ✓ „Übernehmen" (`accept`), ✕ „Ablehnen"
(`reject`), ✎ „Wert eintragen" (`fill`), ✦ „Agent fragen" (`ask`), ↺ „KI-Wert wiederherstellen"
(`restore`), each a 24px button with that `title` and `aria-label`. A press posts `vound:fieldAction
{ id, action }` and hides the chip; a button whose action the field no longer offers (the markers changed
under the open chip) posts nothing. A block holding several fields gets one row per field, labelled with
`SlotMarker.label`, in the order of its `data-vound-slot` ids; a single field's row has no label. Fields
without actions get no row, and a block whose fields have none gets no chip. A press does not take the
focus from the document.

The chip sits in the gutter 10px right of the page (`.doc-page`), vertically centred on the block, in a
zero-height layer appended to the scroll container (`.editor-scroll`), so it scrolls with the page. A dashed
1px tether runs at the block's height from where the page-level block around it ends (the paragraph, or
the whole table for a cell) to the chip, across the page's right margin only, so neither covers text. The
chip stays while the pointer moves from the block to it and hides 220ms after the pointer left both; a
scroll, a window resize or leaving the window hides it too. In a multi-column section the tether is left
out, because the next column may lie between the block and the margin.

The chip never grows past the gutter (the scroll container's right edge without the scrollbar, minus the
page's right edge, less the 10px distance and 16px of slack): its `max-width` is set from it, and a
field's label that does not fit beside its buttons takes a line of its own above them and is cut short,
with the full label in its `title`. Only a gutter narrower than the buttons hides the chip; the host's bar
offers the current field's buttons then. Review mode keeps that gutter free (`reviewGutter.ts`):
`host.css` centres the page in a three-column grid whose right column never gets less than
`--vound-gutter` (`GUTTER_RESERVE_PX`), so a narrow canvas moves the page left, to the canvas edge at the
width fit, and the `reviewFitZoom` transform in `vite.config.ts` makes App.tsx's width
fit (`rawFitWidth`, the overflow clamp, and `fitZoomFor`, Page Width / Text Width / One Page and the fit
mode that follows resizes) measure the canvas through the shim's `reviewFitWidth`, which subtracts
`GUTTER_RESERVE_PX` (110px: a three-button chip, its distance and its slack), and cap the fit at 100 % through
`reviewFitZoom`. A canvas narrower than the page plus the reserve zooms the page out (900px: 93 % for A4),
also when the app locked into fit mode at mount; a wide one keeps 100 %. A zoom the reviewer picks by hand
is not capped; above the fit the chip may not show. Outside review mode the shim returns width and zoom
unchanged.

The hovered block gets `vound-slot-hot`, a 1px ring at 45 %, through a decoration in the markers plugin
(`setHot`, mapped through edits): a class written onto the block's DOM would read as a DOM change to
ProseMirror. In review mode `host.css` restyles the markers: a soft tint and a 2px leading bar (a
background layer, so no box changes), no outline. Proposed is violet, a field still needing a value (gap,
conditional with a dashed bar, refused) amber, accepted cyan with a ✓ after its text (inline with a
negative margin, so it never adds a line; an accepted field emptied again shows none, and a field over several
blocks shows it after the last one, `vound-slot-continued` marking the others), rejected white with a rose
bar and a 1px rose edge. A marked paragraph whose neighbour right above or below is marked for other fields
(`vound-slot-other-above`, `vound-slot-other-below`, set in `markers.ts`) stops its tint and bar 1px short of that
edge, so two adjacent fields show a 2px gap and each its own rounded ends; blocks of one field still join. The current
field has a 1.5px ring and a 5px halo, violet, amber for a field still needing a value. Every colour is a
custom property under `html.vound-review`. Without review mode the markers keep the PR #3362 look.

The current and the hover ring are not drawn on the blocks (`fieldRing.ts`): a field can span consecutive
blocks (paragraphs, list items, or the cells of one column in consecutive rows, all carrying the same
`data-vound-slot`), and a ring on each block either ran a line between them or lost its bottom under the
tint of the marked block after it, which paints later. Instead each run gets one box, its blocks' rects
joined and widened by how far their tint reaches (`--vs-l`, `--vs-r`, `--vs-y`), in a zero-size layer at the
end of the scroll container, above the page and outside layout. Hovering one block of a run rings the whole
run. The layer redraws after every editor update, on a resize of the editor root and of the window.

## Field writes

`vound:setFieldValue` (`fieldWrite.ts`) writes exactly or refuses. The block is the paragraph whose source
index the `BlockRef` names, or the table cell around it; an index that names no node or several apart is `noBlock`;
of a split paragraph only the first half carries the index, so a write or caret finds that half. The block's visible text is a paragraph's text, a field passthrough paragraph's
result text (`fieldDisplay.left`), and for a cell its paragraphs' texts joined with line feeds.

- `previous` occurs exactly once: that occurrence is replaced by `value` (empty allowed). In a paragraph
  it is a text replacement in the formatting of the replaced text; inline leaves (breaks, images) at its
  edges stay. In a field passthrough it rewrites `left` and drops the stale formatted `runs`, as the
  editor's own field editing does; the save writes it into the field's result.
- A field passthrough shows its result trimmed, and an empty legacy field's result is whitespace only (en
  spaces, NBSP): when `previous` does not occur as sent, its trimmed text is looked for in the block's field
  results, and a whitespace-only or empty `previous` matches the one empty field result. An emptied field
  shows as its empty line, so it can be typed into and written again. A paragraph keeps the whitespace of
  a field result as text, so there `previous` matches as sent.
- `previous` is empty and the block is a cell with no text: `value` goes into its first paragraph.
- Otherwise the document is untouched: `absent` (no occurrence, or an empty `previous` in a paragraph),
  `ambiguous` (several, overlapping ones included), `notEditable` (the occurrence spans two paragraphs of
  a cell, or lies in a table-of-contents line, which this rule leaves alone).

A written place is answered with its gap, `{ before, after }`: the block's visible text right before and
right after the value, so a field emptied by a reject can be written again where it was, also after a save
and reload. Both sides grow a character at a time, from nothing up to 32 characters, until the gap names
that place alone; `before: ''` anchors at the start of the block's text, `after: ''` at its end. With no
unique gap within 32 characters the answer has none. A write with `replaces: 'gap'` replaces what lies
after `before` up to the first `after` from there, nothing included, in the formatting of the text before
it; it is refused as `absent` when the anchors no longer occur in that order (the reviewer changed the
text around the field), `ambiguous` when they name several places, and `notEditable` when the place
spans two paragraphs of a cell. Text writes are unchanged and answer a gap too.

Every write is one transaction with the meta `voundHostWrite` (`HOST_WRITE_META`), so it is undoable
and marks the document dirty like typing, but is never reported as `vound:blockEdited`.
`vound:placeCaret` puts a text caret at the end of a paragraph (a cell's last paragraph), focuses the
editor and scrolls the caret into view. For a field passthrough it enters the field the way a double
click does: the DOM caret goes to the end of the result, and ProseMirror's selection moves next to the
block, never a selection of the block itself. A blank result (whitespace only, such as Word's five en
spaces) and an empty result's placeholder are selected instead, in a paragraph as in a field passthrough,
so typing replaces them. Writing five en spaces (Word's empty result) is an ordinary replace, and the en
spaces are then the `previous` of the next write.

## Develop

```sh
npx vite --config web-full/vite.config.ts    # http://localhost:5201 (standalone mode with a sample panel)
```

Standalone mode serves `web-full/samples/*.docx` (gitignored), answers `fontMetrics` from this
machine's installed fonts through `@genoffice/font-metrics` (the dev-only `/__font-metrics` endpoint)
and writes saves to `GENOFFICE_SAVE_DIR`.

## Font metrics

The production build has no font provider: `shim.ts` leaves `window.desktop.fontMetrics` undefined
outside dev, and the renderer (`apps/docs/src/renderer/font-resolve.ts`) treats that as having no
installed-font index. It then lays out from the data it bundles anyway, `word-faces.json`: the macOS
snapshot (`system`) where Word for Mac uses the installed copy, Word's own face table (`faces`) for
everything else. The editor therefore measures text as Word for Mac does on a stock Mac, the same for
every reader. Times New Roman comes from the macOS copy (line factor 1.1499; Word's own copy is
1.1074); Arial, Calibri, Cambria and Verdana from Word's table.
`__tests__/fontResolution.test.ts` pins these answers.

The dev server answers from the machine's own fonts instead. On a Mac without Office that comes out
close to the snapshot; on another machine it may not.
