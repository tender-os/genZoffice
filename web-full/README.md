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
| host → editor | `{ type: 'vound:open', name, bytes: ArrayBuffer, mode?: 'review' }` open a document (transfer the buffer); `mode: 'review'` opens it in review mode, any other mode or none in the normal editor |
| host → editor | `{ type: 'vound:command', command: 'save' }` save now, the same path as ⌘S |
| host → editor | `{ type: 'vound:markers', markers: SlotMarker[] }` replace all slot markers (`{ id, paragraph, scope: 'paragraph' \| 'cell', tone, label, actions }`; `actions` ⊂ `accept`, `reject`, `fill`, `ask`, `restore`; without them a marker gets `label: id`, `actions: []`) |
| host → editor | `{ type: 'vound:scrollTo', id }` scroll a marker to the middle of the view and make it current |
| host → editor | `{ type: 'vound:focus', id: string \| null }` make a marker current without scrolling, or none |
| host → editor | `{ type: 'vound:setFieldValue', requestId, block: BlockRef, previous, value }` replace a field's value in one block (`BlockRef = { paragraph, scope: 'paragraph' \| 'cell' }`, named like a marker); answered with `vound:fieldValueResult` (see Field writes) |
| host → editor | `{ type: 'vound:placeCaret', block: BlockRef }` put the caret at the end of the block's text, focus the editor and scroll it into view |
| editor → host | `{ type: 'vound:ready' }` after the editor subscribed to open requests |
| editor → host | `{ type: 'vound:saved', name, bytes: ArrayBuffer }` on every save (⌘S, File › Save, the save command) |
| editor → host | `{ type: 'vound:dirty', dirty: boolean }` when unsaved changes appear or a save clears them |
| editor → host | `{ type: 'vound:error', message }` open or save failed |
| editor → host | `{ type: 'vound:fieldAction', id, action }` a field chip's button was pressed (`id` is the slot id) |
| editor → host | `{ type: 'vound:blockEdited', ids }` a user edit changed the text of a marked block; `ids` are its slot ids. Review mode only; host writes are not reported |
| editor → host | `{ type: 'vound:caretBlock', ids, leftWith? }` the caret entered a marked block, or left one (`ids: []`, `leftWith: 'Tab' \| 'other'`). Review mode only |
| editor → host | `{ type: 'vound:fieldValueResult', requestId, ok, reason? }` the answer to `vound:setFieldValue` (`reason`: `absent`, `ambiguous`, `notEditable`, `noBlock`) |

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
is. An index that matches no node or more than one (a split copies the index to both
halves) is dropped and logged with `console.info('[vound] markers dropped', ids)`; so is a cell marker
outside a table cell. Paragraphs of nested tables are not editor nodes and cannot be marked.

`markersPlugin.ts` keeps the decorations through edits and walks the document again when the block
count changes (phased loading appends blocks after the first 64, splits, deletions), when an attribute
change dropped a decoration, or while markers wait for their document. The wait ends at the first
build that sees nodes with source indices, so markers that all miss a loaded document do not cost a
walk per edit. Markers sent before the app handed over its editor (`hostEditorHandle` in
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
from the AI always, and in review mode formatted Paste, the field commands and New Comment.

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
negative margin, so it never adds a line), rejected white with a rose bar and a 1px rose edge. The current
field has a 1.5px ring and a 5px halo, violet, amber for a field still needing a value. Every colour is a
custom property under `html.vound-review`. Without review mode the markers keep the PR #3362 look.

## Field writes

`vound:setFieldValue` (`fieldWrite.ts`) writes exactly or refuses. The block is the paragraph whose source
index the `BlockRef` names, or the table cell around it; an index that names no node or several (a split
paragraph) is `noBlock`. The block's visible text is a paragraph's text, a field passthrough paragraph's
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
