/**
 * Placeholder text (Vound fork): a run the engine marks as a content control's placeholder or a form field's
 * default text carries a ctrlPlaceholder mark into the editor and back, and typed text never inherits it.
 */
import { Editor } from '@tiptap/core'
import { afterEach, describe, expect, it } from 'vitest'
import { inlineToRuns, runsToInline } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

describe('placeholder text', () => {
  it('rides a ctrlPlaceholder mark into the editor and back', () => {
    const inline = runsToInline([{ text: 'Firma: ' }, { text: 'Name eingeben', placeholder: true }])
    expect(inline.map((node) => (node.marks ?? []).map((mark) => mark.type))).toEqual([[], ['ctrlPlaceholder']])
    expect(inlineToRuns(inline)).toEqual([{ text: 'Firma: ' }, { text: 'Name eingeben', placeholder: true }])
  })

  it('is a mark of the editor schema that typed text does not continue', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [{ type: 'docParagraph', content: runsToInline([{ text: 'Rechtsform', placeholder: true }]) }] } as never,
    })
    editors.push(editor)
    editor.commands.setTextSelection({ from: 1, to: 11 })
    editor.commands.insertContent('GmbH')
    expect(editor.state.doc.textContent).toBe('GmbH')
    expect(editor.state.doc.firstChild!.firstChild!.marks.map((mark) => mark.type.name)).toEqual([])
  })
})
