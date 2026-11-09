/** In review mode a field the editor shows as a chip takes no text: typing, pasting or deleting on the selected
 * chip leaves the form as it was, in either edit scope (web-full/editScope.ts). Without that, a keystroke on
 * the chip replaced its whole paragraph and the save rewrote the form's own text. */
import { Editor } from '@tiptap/core'
import { Slice } from '@tiptap/pm/model'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editScopePlugin, setEditScope } from '../../../web-full/editScope'
import { markersPlugin, setMarkers } from '../../../web-full/markersPlugin'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

// A blank legacy text form field indented by a tab: it stays a chip, since a save could not tell the tab's
// line from the field's result (parse-fields.ts formTextResults).
const TABBED_BLANK_FIELD =
  '<w:p><w:r><w:tab/></w:r>' +
  '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Text1"/><w:enabled/><w:textInput/></w:ffData></w:fldChar></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>     </w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
const BODY =
  '<w:p><w:r><w:t>Frau/Herr</w:t></w:r></w:p>' +
  TABBED_BLANK_FIELD +
  '<w:p><w:r><w:t>ist heute verpflichtet worden.</w:t></w:r></w:p>'

const documentXmlOf = async (bytes: Uint8Array) =>
  (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')

let editor: Editor | null = null
afterEach(() => {
  editor?.destroy()
  editor = null
  setEditScope('everything')
})

async function openInReview(scope: 'fields' | 'everything') {
  const bytes = await buildDocx({ bodyXml: BODY })
  const parsed = await parseDocx(bytes)
  editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  const post = vi.fn()
  editor.registerPlugin(markersPlugin)
  editor.registerPlugin(editScopePlugin(post))
  setMarkers(editor.view, [
    { id: 's1', paragraph: 1, scope: 'paragraph', tone: 'gap', label: 'Name', actions: ['fill'] },
  ])
  setEditScope(scope)
  let chipPos = -1
  editor.state.doc.forEach((node, pos) => {
    if (node.type.name === 'docProtected') chipPos = pos
  })
  return { bytes, parsed, editor, post, chipPos }
}

const savedXml = async (opened: Awaited<ReturnType<typeof openInReview>>) => {
  const plan = pmDocToSavePlan(opened.editor.getJSON() as PmNode, opened.parsed.blocks)
  return documentXmlOf(await saveDocx(opened.parsed, plan.saveBlocks))
}

describe('a field chip in review mode', () => {
  for (const scope of ['fields', 'everything'] as const) {
    it(`keeps the field and the lines around it when the reviewer types, pastes or deletes on it (${scope})`, async () => {
      const opened = await openInReview(scope)
      const { view } = opened.editor
      expect(view.state.doc.nodeAt(opened.chipPos)?.attrs.fieldDisplay ?? null).toBeNull()
      view.dispatch(
        view.state.tr.setSelection(NodeSelection.create(view.state.doc, opened.chipPos)),
      )

      view.dispatch(view.state.tr.insertText('are be Limonka'))
      const pasted = view.state.schema.nodes.docParagraph.create(
        null,
        view.state.schema.text('Limonka'),
      )
      view.dispatch(
        view.state.tr.replaceSelection(
          new Slice(view.state.schema.nodes.doc.create(null, pasted).content, 0, 0),
        ),
      )
      view.dispatch(view.state.tr.deleteSelection())

      expect(opened.post).toHaveBeenCalledWith({ type: 'vound:editRefused' })
      expect(await savedXml(opened)).toBe(await documentXmlOf(opened.bytes))
    })
  }

  it('keeps the chip when a deletion runs from the text before it over it', async () => {
    const opened = await openInReview('everything')
    const { view } = opened.editor
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 4, opened.chipPos + 3)),
    )
    view.dispatch(view.state.tr.deleteSelection())
    expect(await savedXml(opened)).toBe(await documentXmlOf(opened.bytes))
  })
})
