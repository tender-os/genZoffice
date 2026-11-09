// @vitest-environment jsdom
import { Editor } from '@tiptap/core'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { AllSelection, TextSelection } from '@tiptap/pm/state'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../../apps/docs/src/renderer/editor/convert'
import { editorExtensions } from '../../apps/docs/src/renderer/editor/extensions'
import { editScopePlugin, setEditScope } from '../editScope'
import { markersPlugin, setMarkers } from '../markersPlugin'

const para = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
const BODY = `${para('Kopf des Formulars')}${para('Erstes Feld')}${para('Geschuetzter Text')}${para('Zweites Feld')}${para('Fuss')}`
const TEXTS = ['Kopf des Formulars', 'Erstes Feld', 'Geschuetzter Text', 'Zweites Feld', 'Fuss']

const editors: Editor[] = []
afterEach(() => {
  editors.splice(0).forEach((editor) => editor.destroy())
  setEditScope('everything')
})

async function open(scope: 'fields' | 'everything') {
  const parsed = await parseDocx(await buildDocx({ bodyXml: BODY }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  editor.registerPlugin(markersPlugin)
  editor.registerPlugin(editScopePlugin(vi.fn()), (plugin, plugins) => [plugin, ...plugins])
  const indexOf = (position: number) => editor.state.doc.child(position).attrs.docxPara
  setMarkers(editor.view, [
    {
      id: 'a',
      paragraph: indexOf(1),
      scope: 'paragraph',
      tone: 'gap',
      label: 'A',
      actions: ['ask'],
    },
    {
      id: 'b',
      paragraph: indexOf(3),
      scope: 'paragraph',
      tone: 'gap',
      label: 'B',
      actions: ['ask'],
    },
  ])
  setEditScope(scope)
  editors.push(editor)
  return editor
}

const texts = (editor: Editor) => {
  const out: string[] = []
  editor.state.doc.forEach((node) => out.push(node.textContent))
  return out
}

const posIn = (editor: Editor, text: string) => {
  let found = -1
  editor.state.doc.descendants((node, pos) => {
    if (found === -1 && node.isText && node.text?.includes(text))
      found = pos + node.text.indexOf(text)
  })
  return found
}

const pressSelectAll = (editor: Editor) =>
  editor.view.someProp('handleKeyDown', (handler) =>
    handler(editor.view, new KeyboardEvent('keydown', { key: 'a', metaKey: true })),
  ) ?? false

const selectEverything = (editor: Editor) =>
  editor.view.dispatch(editor.state.tr.setSelection(new AllSelection(editor.state.doc)))

const selectTexts = (editor: Editor, from: string, to: string) =>
  editor.view.dispatch(
    editor.state.tr.setSelection(
      TextSelection.create(editor.state.doc, posIn(editor, from), posIn(editor, to) + to.length),
    ),
  )

const EDITS: Array<[string, (editor: Editor) => void]> = [
  [
    'Backspace',
    (editor) => {
      editor.commands.keyboardShortcut('Backspace')
    },
  ],
  [
    'Delete',
    (editor) => {
      editor.commands.keyboardShortcut('Delete')
    },
  ],
  ['delete selection', (editor) => editor.view.dispatch(editor.state.tr.deleteSelection())],
  ['typing', (editor) => editor.view.dispatch(editor.state.tr.insertText('x'))],
  [
    'a cut',
    (editor) => editor.view.dispatch(editor.state.tr.deleteSelection().setMeta('uiEvent', 'cut')),
  ],
  [
    'a paste',
    (editor) =>
      editor.view.dispatch(
        editor.state.tr
          .replaceSelectionWith(editor.schema.text('eingefuegt'))
          .setMeta('paste', true)
          .setMeta('uiEvent', 'paste'),
      ),
  ],
]

describe('locked to the fields', () => {
  it('select-all in a field selects only that field', async () => {
    const editor = await open('fields')
    editor.commands.setTextSelection(posIn(editor, 'Erstes') + 2)
    expect(pressSelectAll(editor)).toBe(true)
    const { from, to } = editor.state.selection
    expect(editor.state.doc.textBetween(from, to)).toBe('Erstes Feld')
  })

  it('select-all outside a field changes nothing', async () => {
    const editor = await open('fields')
    editor.commands.setTextSelection(posIn(editor, 'Kopf') + 1)
    const before = editor.state.selection.from
    expect(pressSelectAll(editor)).toBe(true)
    expect(editor.state.selection.from).toBe(before)
    expect(editor.state.selection.empty).toBe(true)
  })

  it.each(EDITS)('a selection of the whole document survives %s', async (_, edit) => {
    const editor = await open('fields')
    selectEverything(editor)
    edit(editor)
    expect(texts(editor)).toEqual(TEXTS)
  })

  it.each(EDITS)('a selection from first to last character survives %s', async (_, edit) => {
    const editor = await open('fields')
    selectTexts(editor, 'Kopf', 'Fuss')
    edit(editor)
    expect(texts(editor)).toEqual(TEXTS)
  })

  it.each(EDITS)(
    'a selection across two fields and the text between survives %s',
    async (_, edit) => {
      const editor = await open('fields')
      selectTexts(editor, 'Erstes', 'Zweites Feld')
      edit(editor)
      expect(texts(editor)).toEqual(TEXTS)
    },
  )

  it('a selection inside one field can still be replaced', async () => {
    const editor = await open('fields')
    selectTexts(editor, 'Erstes', 'Feld')
    editor.view.dispatch(editor.state.tr.insertText('Neu'))
    expect(texts(editor)[1]).toBe('Neu')
  })

  it('a document opening still replaces everything', async () => {
    const editor = await open('fields')
    editor.commands.setContent({
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'Neu' }] }],
    })
    expect(texts(editor)).toEqual(['Neu'])
  })
})

describe('editing the whole document', () => {
  it("select-all stays the editor's own, and deleting everything works", async () => {
    const editor = await open('everything')
    editor.commands.setTextSelection(posIn(editor, 'Erstes') + 2)
    expect(pressSelectAll(editor)).toBe(false)
    selectEverything(editor)
    editor.view.dispatch(editor.state.tr.deleteSelection())
    expect(editor.state.doc.textContent).toBe('')
  })
})
