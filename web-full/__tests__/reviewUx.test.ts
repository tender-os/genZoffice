// @vitest-environment jsdom
import { Editor } from '@tiptap/core'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { TextSelection } from '@tiptap/pm/state'
import { EMPTY_CONTROL_TEXT, parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../../apps/docs/src/renderer/editor/convert'
import { editorExtensions } from '../../apps/docs/src/renderer/editor/extensions'
import { editScopePlugin, setEditScope } from '../editScope'
import { readHostMessage } from '../embedMessages'
import { formTextPlugin } from '../formTextField'
import { historyReportPlugin, MENU_COMMANDS, refocusOnReturn } from '../hostCommands'
import { markersPlugin, setCurrent, setMarkers } from '../markersPlugin'
import { EMPTY_CONTROL_CLASS, reviewModePlugin } from '../reviewMode'

const plain = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const field = (text: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="F"/><w:enabled/><w:textInput/></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${plain(text)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`
const control = (text: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="5"/><w:text/></w:sdtPr><w:sdtContent>${plain(text)}</w:sdtContent></w:sdt>`
const BODY = [
  `<w:p>${plain('Vertretung: ')}${field('Prokurist')}</w:p>`,
  `<w:p>${plain('Fester Text des Formulars')}</w:p>`,
  `<w:p>${plain('Ort: ')}${control('Leipzig')}</w:p>`,
].join('')

const editors: Editor[] = []
afterEach(() => {
  editors.splice(0).forEach((editor) => editor.destroy())
  document.body.replaceChildren()
  setEditScope('everything')
})

async function open(post = vi.fn(), scope: 'fields' | 'everything' = 'fields') {
  const parsed = await parseDocx(await buildDocx({ bodyXml: BODY }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  // as the app opens a document: not an edit to undo
  editor
    .chain()
    .setMeta('addToHistory', false)
    .setContent(blocksToPmDoc(parsed.blocks) as never)
    .run()
  const first = <T>(plugin: T, plugins: T[]) => [plugin, ...plugins]
  editor.registerPlugin(markersPlugin)
  editor.registerPlugin(reviewModePlugin, first)
  editor.registerPlugin(formTextPlugin, first)
  editor.registerPlugin(editScopePlugin(post), first)
  document.body.append(editor.view.dom)
  setMarkers(
    editor.view,
    [0, 2].map((index) => ({
      id: `m${index}`,
      paragraph: editor.state.doc.child(index).attrs.docxPara,
      scope: 'paragraph' as const,
      tone: 'gap' as const,
      label: 'F',
      actions: ['ask' as const],
    })),
  )
  setEditScope(scope)
  editors.push(editor)
  return editor
}

const posOf = (editor: Editor, text: string) => {
  let found = -1
  editor.state.doc.descendants((node, pos) => {
    if (found === -1 && node.isText && node.text?.includes(text))
      found = pos + node.text.indexOf(text)
  })
  return found
}

const textNodeOf = (editor: Editor, text: string) => {
  const walker = document.createTreeWalker(editor.view.dom, NodeFilter.SHOW_TEXT)
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (node.textContent?.includes(text)) return node
  }
  throw new Error(`no text node with ${text}`)
}

/** the browser's typing: the view's text input handlers, else an insertion at the selection */
function type(editor: Editor, text: string) {
  const { view } = editor
  const { from, to } = view.state.selection
  const handled = view.someProp('handleTextInput', (handler) =>
    handler(view, from, to, text, () => view.state.tr.insertText(text, from, to)),
  )
  if (!handled) view.dispatch(view.state.tr.insertText(text, from, to))
}

const settle = () => new Promise((done) => setTimeout(done, 60))

describe('a click on locked text while the caret is in a field', () => {
  it('moves the caret there, so typing is refused and never lands in the field', async () => {
    const post = vi.fn()
    const editor = await open(post)
    editor.view.focus()
    const end = posOf(editor, 'Prokurist') + 'Prokurist'.length
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, end)))

    // the click puts the DOM caret into the label; a marker update lands before the editor reads it
    document.getSelection()!.collapse(textNodeOf(editor, 'Fester'), 3)
    setCurrent(editor.view, 'm0')
    document.dispatchEvent(new Event('selectionchange'))
    await settle()

    expect(editor.state.selection.from).toBe(posOf(editor, 'Fester') + 3)
    type(editor, 'Z')
    expect(editor.state.doc.child(0).textContent).toBe('Vertretung: Prokurist')
    expect(editor.state.doc.child(1).textContent).toBe('Fester Text des Formulars')
    expect(post).toHaveBeenCalledWith({ type: 'vound:editRefused' })
  })
})

describe('text the browser writes into locked text (a dead key, an input method, autofill)', () => {
  it('is refused and taken out of the page again', async () => {
    const post = vi.fn()
    const editor = await open(post)
    const label = textNodeOf(editor, 'Fester')
    label.textContent = 'Fest´er Text des Formulars'
    await settle()
    expect(editor.state.doc.child(1).textContent).toBe('Fester Text des Formulars')
    expect(editor.view.dom.textContent).not.toContain('´')
    expect(post).toHaveBeenCalledWith({ type: 'vound:editRefused' })
  })
})

describe('undo after a field was emptied', () => {
  it('brings a legacy text field back without its filler', async () => {
    const editor = await open()
    const from = posOf(editor, 'Prokurist')
    editor.view.dispatch(editor.state.tr.delete(from, from + 'Prokurist'.length))
    expect(editor.state.doc.child(0).textContent).toContain(EMPTY_CONTROL_TEXT)
    editor.commands.undo()
    expect(editor.state.doc.child(0).textContent).toBe('Vertretung: Prokurist')
    expect(editor.view.dom.querySelector(`.${EMPTY_CONTROL_CLASS}`)).toBeNull()
    editor.commands.redo()
    expect(editor.state.doc.child(0).textContent).toBe(`Vertretung: ${EMPTY_CONTROL_TEXT}`)
  })

  it('brings a content control back without the filler a save gave the emptied control', async () => {
    const post = vi.fn()
    const editor = await open(post, 'everything')
    const from = posOf(editor, 'Leipzig')
    editor.view.dispatch(editor.state.tr.delete(from, from + 'Leipzig'.length))
    // the app takes over the saved file's parse: the emptied control shows its filler, outside the history
    const filler = editor.schema.text(EMPTY_CONTROL_TEXT, [
      editor.schema.marks.ctrlPlaceholder.create(),
    ])
    editor.view.dispatch(editor.state.tr.insert(from, filler).setMeta('addToHistory', false))
    editor.commands.undo()
    expect(editor.state.doc.child(2).textContent).toBe('Ort: Leipzig')
    expect(editor.view.dom.querySelector(`.${EMPTY_CONTROL_CLASS}`)).toBeNull()
  })
})

describe('undo, redo and zoom from the host', () => {
  it('reads the commands and the zoom the host may send', () => {
    for (const command of ['undo', 'redo', 'zoomIn', 'zoomOut', 'zoomFitWidth'] as const) {
      expect(readHostMessage({ type: 'vound:command', command })).toEqual({
        type: 'vound:command',
        command,
      })
    }
    expect(readHostMessage({ type: 'vound:setZoom', zoom: 125 })).toEqual({
      type: 'vound:setZoom',
      zoom: 125,
    })
    expect(readHostMessage({ type: 'vound:setZoom', zoom: '125' })).toBeNull()
    expect(readHostMessage({ type: 'vound:setZoom', zoom: Number.NaN })).toBeNull()
  })

  it("runs each command through a command the app's menu handles", () => {
    const app = readFileSync(resolve(__dirname, '../../apps/docs/src/renderer/App.tsx'), 'utf8')
    for (const command of Object.values(MENU_COMMANDS)) expect(app).toContain(`case '${command}':`)
  })

  it('reports whether there is something to undo or redo, once per change', async () => {
    const post = vi.fn()
    const editor = await open(vi.fn(), 'everything')
    editor.registerPlugin(historyReportPlugin(post))
    expect(post).toHaveBeenLastCalledWith({ type: 'vound:history', canUndo: false, canRedo: false })
    type(editor, 'x')
    expect(post).toHaveBeenLastCalledWith({ type: 'vound:history', canUndo: true, canRedo: false })
    const reports = post.mock.calls.length
    type(editor, 'y')
    expect(post.mock.calls.length).toBe(reports)
    editor.commands.undo()
    expect(post).toHaveBeenLastCalledWith({ type: 'vound:history', canUndo: false, canRedo: true })
  })

  it('gives the editor its focus back when the host returns focus to the frame', async () => {
    const editor = await open()
    const stop = refocusOnReturn(() => editor.view)
    editor.view.dom.blur()
    expect(editor.view.hasFocus()).toBe(false)
    window.dispatchEvent(new Event('focus'))
    expect(editor.view.hasFocus()).toBe(true)
    stop()
  })
})

describe('an accepted field', () => {
  it('takes the success colour, as the legend and the PDF review do', () => {
    const css = readFileSync(resolve(__dirname, '../host.css'), 'utf8')
    expect(css).toMatch(/--vound-accepted-bar: hsl\(152 55% 38%\)/)
    expect(css).toMatch(/--vound-accepted-tint: hsl\(152 /)
  })
})
