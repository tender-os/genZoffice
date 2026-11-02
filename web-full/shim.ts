// Browser stand-in for the Electron preload bridge (window.desktop) of the Docs app.
// Open: a browser file picker (File > Open) or a sample from /samples (?doc=name, or the host panel).
// Save: every save is POSTed to the dev server (GENOFFICE_SAVE_DIR) and offered as a browser download.
// Font metrics: the dev server answers from this machine's installed fonts. The production build has no
// provider, so the renderer lays out from its bundled macOS snapshot (font-resolve.ts treats a missing
// window.desktop.fontMetrics as no installed-font index).
// Markers: the app hands over its editor (vite.config.ts, hostEditorHandle); the shim registers the
// markers plugin on it and routes vound:markers, vound:scrollTo and vound:focus to it.
// Review mode: vound:open with mode 'review' puts the handed-over editor in review mode (reviewMode.ts),
// which reports edits and caret moves in marked blocks and shows the field chip; an open without it leaves
// review mode. A save in review mode does not seal the markers (markersPlugin.ts, markersOnSave).
// Field writes: vound:setFieldValue and vound:placeCaret act on the handed-over editor (fieldWrite.ts); every
// setFieldValue is answered with vound:fieldValueResult.
// Language: an embedding host picks the UI language with ?lang= (embedLocale.ts); stand-alone stays English.
// Zoom: the app's width fit asks reviewFitWidth and reviewFitZoom (the reviewFitZoom transform in
// vite.config.ts); in review mode they keep the field chip's gutter free and cap the fit at 100 %.
import type { Editor } from '@tiptap/core'
import { htmlLang } from '@genoffice/i18n'
import { isEmbedded, type EditorToHost, type SlotMarker } from './embed-protocol'
import { acceptsHostMessage, readHostMessage } from './embedMessages'
import { embeddedSave } from './embeddedSave'
import { embedLang } from './embedLocale'
import { applyFieldValue, placeCaret } from './fieldWrite'
import { markersKey, markersOnSave, markersPlugin, scrollToMarker, setCurrent, setMarkers, resetMarkers } from './markersPlugin'
import { enterReviewMode, leaveReviewMode } from './reviewMode'
import { reviewFitWidth, reviewFitZoom } from './reviewGutter'
import { menuShows } from './reviewMenu'
type Handler = (...args: unknown[]) => void
const calls: string[] = []
let openHandler: Handler | null = null
let menuHandler: Handler | null = null
let hostEditor: Editor | null = null
/** markers that arrived before the app handed over its editor */
let pendingMarkers: SlotMarker[] | null = null
/** the mode of the last vound:open; applied when the app hands over its editor if the open came first */
let reviewMode = false
const embedded = isEmbedded()
// same-origin only: the host and the editor are served from one origin (embed-protocol.ts)
const toHost = (msg: EditorToHost, transfer: Transferable[] = []) => { if (embedded) window.parent.postMessage(msg, window.location.origin, transfer) }
const applyMode = (editor: Editor) => (reviewMode ? enterReviewMode(editor, toHost) : leaveReviewMode(editor))
/** the answer to a field write that arrives before the app handed over its editor */
const noEditor = { ok: false, reason: 'noBlock' } as const
// the app reports its unsaved-changes state on every render (vite.config.ts, hostReportsDirty); the host hears changes only
let lastDirty = false
const hostDirty = (dirty: boolean) => { if (dirty === lastDirty) return; lastDirty = dirty; toHost({ type: 'vound:dirty', dirty }) }
const params = new URLSearchParams(location.search)
const docName = params.get('doc')
const uiLang = embedLang(location.search, embedded)
const lastSave: { name: string; bytes: Uint8Array | null; url: string | null } = { name: '', bytes: null, url: null }
async function devFontMetrics(family: string) {
  const res = await fetch(`/__font-metrics?family=${encodeURIComponent(family)}`)
  return res.ok ? res.json() : null
}
const sampleResult = (name: string) => ({ path: `/samples/${name}.docx`, name: `${name}.docx`, dataUrl: `/samples/${name}.docx`, hash: '' })
const fileResult = (file: File) => ({ path: `/local/${file.name}`, name: file.name, dataUrl: URL.createObjectURL(file), hash: '' })
function pickFile(): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input'); input.type = 'file'; input.accept = '.docx'
    input.onchange = () => resolve(input.files?.[0] ?? null)
    // no change event fires on cancel in every browser; a focus return without a file counts as cancel
    window.addEventListener('focus', () => setTimeout(() => resolve(input.files?.[0] ?? null), 400), { once: true })
    input.click()
  })
}
function download(name: string, bytes: Uint8Array) {
  if (lastSave.url) URL.revokeObjectURL(lastSave.url)
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }))
  lastSave.url = url
  const a = document.createElement('a'); a.href = url; a.download = name.endsWith('.docx') ? name : `${name}.docx`; a.click()
}
async function persist(path: string, data: ArrayBuffer, triggerDownload: boolean, auto?: boolean) {
  // the host owns persistence when embedded; only explicit saves reach it (embeddedSave.ts)
  // review mode keeps the markers through a save (the re-parse keeps only those posted since), the plain embed seals them
  if (embedded) return embeddedSave(path, data, auto, toHost, () => { if (hostEditor) markersOnSave(hostEditor.view, reviewMode ? 'review' : 'plain') })
  const name = path.split('/').pop()!.replace(/\.docx$/, '')
  lastSave.name = name; lastSave.bytes = new Uint8Array(data)
  const res = await fetch(`/__save?name=${encodeURIComponent(name)}`, { method: 'POST', body: data }).catch(() => null)
  if (triggerDownload) download(name, lastSave.bytes)
  return { ok: !!res && res.ok, path }
}
function handOverEditor(editor: Editor) {
  hostEditor = editor
  // registering a keyed plugin twice throws; a re-run effect hands over the same editor again
  if (!markersKey.getState(editor.state)) editor.registerPlugin(markersPlugin)
  if (pendingMarkers) setMarkers(editor.view, pendingMarkers)
  pendingMarkers = null
  applyMode(editor)
}
const known: Record<string, unknown> = {
  getLanguage: async () => uiLang,
  getTheme: async () => 'light',
  getDocumentTheme: async () => 'follow',
  getSystemLocale: async () => htmlLang(uiLang),
  getAuthorName: async () => 'Reviewer',
  getAiPanelPrefs: async () => ({}),
  getAiSettings: async () => ({}),
  getRecentFiles: async () => [],
  getAutoSaveDefault: async () => false,
  args: async () => [],
  cwd: async () => '/',
  allowedRoots: async () => [],
  listPrinters: async () => [],
  kbList: async () => [],
  listDocsTabs: async () => [],
  consumePendingOpenDocx: async () => (docName ? sampleResult(docName) : null),
  consumeNewBlankDoc: async () => !docName,
  consumeAiDocContent: async () => null,
  consumeHeadlessExport: async () => null,
  // undefined, not a stub: the Proxy below would otherwise answer with a function that resolves to nothing
  fontMetrics: import.meta.env.DEV ? devFontMetrics : undefined,
  // File > Open
  openDocx: async () => { const f = await pickFile(); return f ? fileResult(f) : null },
  openDocxPath: async (path: string) => sampleResult(path.split('/').pop()!.replace(/\.docx$/, '')),
  // Save (Cmd+S) writes to the server folder and downloads; Save As / first save of a new doc do the same
  // auto is true for autosave ticks (apps/docs/src/shared/ipc.ts saveDocx)
  saveDocx: (path: string, data: ArrayBuffer, auto?: boolean) => persist(path, data, true, auto),
  saveDocxAs: (defaultName: string, data: ArrayBuffer) => persist(`/local/${defaultName}`, data, true),
  saveDocxNew: (defaultName: string, data: ArrayBuffer) => persist(`/local/${defaultName}`, data, true),
  saveDocxTo: (path: string, data: ArrayBuffer) => persist(path, data, false),
  writeRecoveryCopy: async () => ({ ok: true }),
  onOpenDocx: (h: Handler) => { openHandler = h; if (embedded) toHost({ type: 'vound:ready' }); return () => { openHandler = null } },
  hostDirty,
  onMenuCommand: (h: Handler) => { menuHandler = h; return () => { menuHandler = null } },
  hostEditor: handOverEditor,
  reviewFitWidth: (canvasWidth: number) => reviewFitWidth(canvasWidth, reviewMode),
  reviewFitZoom: (zoom: number) => reviewFitZoom(zoom, reviewMode),
  menuShows: (key: string) => menuShows(key, embedded, reviewMode),
}
export function installDesktopShim() {
  const w = window as unknown as Record<string, unknown>
  w.desktop = new Proxy(known, {
    get(target, key) {
      if (typeof key !== 'string') return undefined
      if (key in target) return target[key]
      if (key.startsWith('on')) return () => () => {}
      return async (...args: unknown[]) => { calls.push(`${key}(${args.length})`); return undefined }
    },
  })
  const api = {
    calls,
    lastSave,
    openSample: (name: string) => { if (!openHandler) throw new Error('App has not subscribed to onOpenDocx yet'); openHandler(sampleResult(name)) },
    openFile: async () => { const f = await pickFile(); if (f && openHandler) openHandler(fileResult(f)) },
    downloadLastSave: () => { if (lastSave.bytes) download(lastSave.name, lastSave.bytes) },
  }
  w.__full = api
  if (embedded) {
    window.addEventListener('message', (event: MessageEvent<unknown>) => {
      if (!acceptsHostMessage(event, { origin: window.location.origin, parent: window.parent })) return
      const msg = readHostMessage(event.data)
      if (!msg) return
      switch (msg.type) {
        case 'vound:open': {
          reviewMode = msg.mode === 'review'
          if (hostEditor) { resetMarkers(hostEditor.view); applyMode(hostEditor) }
          const file = new File([msg.bytes], msg.name.endsWith('.docx') ? msg.name : `${msg.name}.docx`)
          if (openHandler) openHandler(fileResult(file))
          else toHost({ type: 'vound:error', message: 'editor not ready for open yet' })
          break
        }
        case 'vound:command':
          if (menuHandler) menuHandler(msg.command)
          else toHost({ type: 'vound:error', message: 'editor has no command handler yet' })
          break
        case 'vound:markers':
          if (hostEditor) setMarkers(hostEditor.view, msg.markers)
          else pendingMarkers = msg.markers
          break
        case 'vound:scrollTo':
          if (hostEditor) scrollToMarker(hostEditor.view, msg.id)
          break
        case 'vound:focus':
          if (hostEditor) setCurrent(hostEditor.view, msg.id)
          break
        case 'vound:setFieldValue': {
          const result = hostEditor ? applyFieldValue(hostEditor.view, msg) : noEditor
          toHost({ type: 'vound:fieldValueResult', requestId: msg.requestId, ...result })
          break
        }
        case 'vound:placeCaret':
          if (hostEditor) placeCaret(hostEditor.view, msg.block)
          break
      }
    })
  }
  return api
}
