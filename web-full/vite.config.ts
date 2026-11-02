import { resolve } from 'node:path'
import { mkdirSync, writeFileSync } from 'node:fs'
import react from '@vitejs/plugin-react'
import { defineConfig, type Plugin } from 'vite'
import { familyVerticalMetrics } from '../packages/font-metrics/src/index'
import { AI_MENU_KEYS, REVIEW_MENU_KEYS } from './reviewMenu'
const root = resolve(__dirname, '..')
const saveDir = process.env.GENOFFICE_SAVE_DIR ?? resolve(__dirname, 'saved')
function hostEndpoints(): Plugin {
  return {
    name: 'genoffice-host-endpoints',
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        const url = new URL(req.url ?? '/', 'http://localhost')
        if (url.pathname === '/__font-metrics') {
          const family = url.searchParams.get('family') ?? ''
          res.setHeader('content-type', 'application/json')
          res.end(JSON.stringify(familyVerticalMetrics(family)))
          return
        }
        if (url.pathname === '/__save' && req.method === 'POST') {
          const chunks: Buffer[] = []
          req.on('data', (c: Buffer) => chunks.push(c))
          req.on('end', () => {
            mkdirSync(saveDir, { recursive: true })
            const name = (url.searchParams.get('name') ?? 'untitled').replace(/[^\w.-]/g, '_')
            writeFileSync(resolve(saveDir, `${name}.docx`), Buffer.concat(chunks))
            res.setHeader('content-type', 'application/json')
            res.end(JSON.stringify({ ok: true, path: resolve(saveDir, `${name}.docx`) }))
          })
          return
        }
        next()
      })
    },
  }
}

// Spike-only: the host decides whether the body is editable. The desktop app locks the body when the
// document carries an enforced w:documentProtection (edit="forms" / "readOnly"); our review UI fills
// such forms on purpose, so that lock is lifted here. settings.xml keeps the protection untouched on save.
// Tiptap's setEditable emits an 'update' event unless told not to, and the app counts every update as an
// edit. The effect re-runs when isProtected changes, so opening a protected form would report unsaved
// changes (vound:dirty) before the reviewer typed anything. The update is emitted only when editability
// really toggles, because the nested-table node view re-applies contenteditable on 'update'.
function hostDecidesEditability(): Plugin {
  return {
    name: 'genoffice-host-editability',
    transform(code, id) {
      if (!id.endsWith('/apps/docs/src/renderer/App.tsx')) return null
      const needle = 'editor.setEditable(!readMode && !isProtected)'
      if (code.split(needle).length !== 2) throw new Error('host-editability: App.tsx gate not found once, update the spike')
      // the editor's context menu follows the same rule: it opens on a protected form too
      const menuGate = 'if (readMode || isProtected) return'
      if (code.split(menuGate).length !== 2) throw new Error('host-editability: App.tsx context menu gate not found once, update the spike')
      return {
        code: code
          .replace(needle, 'editor.setEditable(!readMode, editor.isEditable !== !readMode)')
          .replace(menuGate, menuGate.replace('readMode || isProtected', 'readMode')),
        map: null,
      }
    },
  }
}

// The host warns before closing an editor with unsaved changes, so it has to hear when they appear and
// when a save clears them. App.tsx recomputes its composite dirty state on every render (every edit path
// re-renders); after each commit, hand it to the shim, which posts vound:dirty when it changed.
const dirtyNeedle = 'anyDirtyRef.current = hasUnsavedChanges'
const dirtyReport = ';\nuseEffect(() => { window.desktop.hostDirty?.(!!doc && hasUnsavedChanges) });\n'
function hostReportsDirty(): Plugin {
  return {
    name: 'genoffice-host-dirty',
    transform(code, id) {
      if (!id.endsWith('/apps/docs/src/renderer/App.tsx')) return null
      if (code.split(dirtyNeedle).length !== 2) throw new Error('host-dirty: App.tsx dirty state not found once, update the transform')
      return { code: code.replace(dirtyNeedle, dirtyNeedle + dirtyReport), map: null }
    },
  }
}

// The markers plugin (markersPlugin.ts) needs the app's editor. App.tsx creates exactly one with
// useEditor; hand it to the shim once it exists, after the dirty report (same needle, appended after it).
function hostEditorHandle(): Plugin {
  return {
    name: 'genoffice-host-editor',
    transform(code, id) {
      if (!id.endsWith('/apps/docs/src/renderer/App.tsx')) return null
      // hostReportsDirty runs first: the needle is followed by its dirty report
      const reported = dirtyNeedle + dirtyReport
      if (code.split(reported).length !== 2) throw new Error('host-editor: App.tsx dirty report not found once, update the transform')
      const handOver = 'useEffect(() => { if (editor) window.desktop.hostEditor?.(editor) }, [editor]);\n'
      return { code: code.replace(reported, reported + handOver), map: null }
    },
  }
}

// Review mode keeps a gutter right of the page for the field chip (reviewGutter.ts): the width fit fills
// the canvas less that gutter and never zooms above 100 %. App.tsx measures the canvas in rawFitWidth
// (the overflow clamp) and fitZoomFor (Page Width / Text Width / One Page, the fit mode that follows
// resizes); both read the canvas width through the shim, and fitZoomFor caps its result. Outside review
// mode the shim answers with the width and zoom unchanged.
const fitCanvas = '(window.desktop.reviewFitWidth?.(scroller.clientWidth) ?? scroller.clientWidth)'
const fitNeedles = [
  { needle: 'return ((scroller.clientWidth - 48) / (', replacement: `return ((${fitCanvas} - 48) / (`, count: 1 },
  { needle: '((scroller.clientWidth - pad) / (', replacement: `((${fitCanvas} - pad) / (`, count: 2 },
  {
    needle: 'return clampDocsZoom(Math.floor(next))',
    replacement: 'return clampDocsZoom(Math.floor(window.desktop.reviewFitZoom?.(next) ?? next))',
    count: 1,
  },
]
function reviewFitZoom(): Plugin {
  return {
    name: 'genoffice-review-fit-zoom',
    // on the source, before esbuild drops the parentheses the needles contain
    enforce: 'pre',
    transform(code, id) {
      if (!id.endsWith('/apps/docs/src/renderer/App.tsx')) return null
      let out = code
      for (const { needle, replacement, count } of fitNeedles) {
        const found = out.split(needle).length - 1
        if (found !== count) throw new Error(`review-fit-zoom: "${needle}" found ${found} times, expected ${count}; update the transform`)
        out = out.split(needle).join(replacement)
      }
      return { code: out, map: null }
    },
  }
}

// The embed leaves some context-menu entries out (reviewMenu.ts): each one renders only when the shim's
// menuShows says so. Outside the embed the shim answers true.
function reviewMenu(): Plugin {
  return {
    name: 'genoffice-review-menu',
    // on the source, before esbuild rewrites the quotes the needles contain
    enforce: 'pre',
    transform(code, id) {
      if (!id.endsWith('/apps/docs/src/renderer/components/ContextMenu.tsx')) return null
      let out = code
      for (const key of [...AI_MENU_KEYS, ...REVIEW_MENU_KEYS]) {
        const needle = `item(t('${key}'),`
        const found = out.split(needle).length - 1
        if (found !== 1) throw new Error(`review-menu: "${needle}" found ${found} times, expected 1; update the transform`)
        out = out.replace(needle, `(window.desktop.menuShows?.('${key}') ?? true) && ${needle}`)
      }
      return { code: out, map: null }
    },
  }
}

// Spike-only fix for a GenOffice rendering bug: header/footer strips sit outside .doc-page, so runs that
// inherit the document's default (theme) font fell back to the UI font stack. Extend the per-document
// base font rule to the strips. Upstream candidate: apps/docs/src/renderer/doc-style-css.ts.
function headerStripFont(): Plugin {
  return {
    name: 'genoffice-header-strip-font',
    transform(code, id) {
      if (!id.endsWith('/apps/docs/src/renderer/doc-style-css.ts')) return null
      const re = /\.doc-page, \.pv-page \{ \$\{decls\.join\(/g
      if (!re.test(code)) throw new Error('header-strip-font: base font rule not found, update the spike')
      return { code: code.replace(re, '.doc-page, .pv-page, .page-hf { ${decls.join('), map: null }
    },
  }
}
export default defineConfig({
  root: __dirname,
  // relative asset URLs so the built host can be served under any path (e.g. /editor/)
  base: './',
  build: { outDir: resolve(__dirname, 'dist'), emptyOutDir: true, target: 'es2022' },
  plugins: [react(), hostEndpoints(), hostDecidesEditability(), hostReportsDirty(), hostEditorHandle(), reviewFitZoom(), reviewMenu(), headerStripFont()],
  resolve: {
    alias: {
      '@genoffice/docx-engine/lazy-media': resolve(root, 'packages/docx-engine/src/lazy-media.ts'),
      '@genoffice/docx-engine/zip-splice': resolve(root, 'packages/docx-engine/src/zip-splice.ts'),
      '@genoffice/docx-engine': resolve(root, 'packages/docx-engine/src/index.ts'),
    },
  },
  server: { port: 5201, strictPort: true, fs: { allow: [root] } },
})
