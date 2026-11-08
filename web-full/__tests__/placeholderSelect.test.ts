// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest'
import { EditorState, TextSelection } from '@tiptap/pm/state'
import { EditorView } from '@tiptap/pm/view'
import { placeholderSelectPlugin } from '../placeholderSelect'
import { doc, fieldPara, para, placeholderPara, schema } from './markerSchema'
import { fieldNodeView } from './fieldNodeView'

// paragraph 0 "Firma: " + placeholder "Name eingeben" + " (GmbH)" at 0..29; text from 1, placeholder 8..21
let view: EditorView | null = null
afterEach(() => { view?.dom.remove(); view?.destroy(); view = null })
const viewOf = () => {
  const state = EditorState.create({ doc: doc(placeholderPara(0, 'Firma: ', 'Name eingeben', ' (GmbH)'), para(1, 'Rechtsform')), plugins: [placeholderSelectPlugin] })
  view = new EditorView(document.createElement('div'), { state })
  return view
}
/** a click at `pos` as ProseMirror reports it to the plugins, then the caret the browser puts there unless one took it */
const click = (v: EditorView, pos: number) => {
  const taken = v.someProp('handleClick', (handle) => handle(v, pos, new MouseEvent('click')))
  if (!taken) v.dispatch(v.state.tr.setSelection(TextSelection.create(v.state.doc, pos)))
  return v
}
/** the caret moved there by a key */
const keyTo = (v: EditorView, pos: number) => {
  v.dispatch(v.state.tr.setSelection(TextSelection.create(v.state.doc, pos)))
  return v
}
const selected = (v: EditorView) => [v.state.selection.from, v.state.selection.to]

describe('placeholder text in review mode', () => {
  it('a click anywhere on placeholder text, its edges included, selects all of it', () => {
    for (const pos of [8, 12, 21]) expect(selected(click(viewOf(), pos))).toEqual([8, 21])
  })

  it('typing then replaces the placeholder whole, and the typed text is no placeholder', () => {
    const v = click(viewOf(), 12)
    v.dispatch(v.state.tr.insertText('Muster GmbH'))
    expect(v.state.doc.firstChild!.textContent).toBe('Firma: Muster GmbH (GmbH)')
    expect(v.state.doc.firstChild!.content.content.flatMap((child) => child.marks)).toEqual([])
  })

  it('a caret moved by keys selects it from inside only, so the reviewer can step past it', () => {
    expect(selected(keyTo(viewOf(), 12))).toEqual([8, 21])
    expect(selected(keyTo(viewOf(), 8))).toEqual([8, 8])
    expect(selected(keyTo(viewOf(), 21))).toEqual([21, 21])
  })

  it('leaves the caret alone in the document text around it and in plain paragraphs', () => {
    for (const pos of [4, 24, 34]) expect(selected(click(viewOf(), pos))).toEqual([pos, pos])
  })
})

describe('a form field outside a table that shows its default text', () => {
  const fieldView = (left: string) => {
    const mount = document.createElement('div')
    document.body.append(mount)
    const state = EditorState.create({ schema, doc: doc(para(0, 'USt-IdNr.:'), fieldPara(1, left, { defaultText: 'USt-IdNr.' })), plugins: [placeholderSelectPlugin] })
    view = new EditorView(mount, { state, nodeViews: { docProtected: fieldNodeView } })
    return view
  }
  const clickField = (v: EditorView) => v.dom.querySelector('.doc-field-text')!.dispatchEvent(new MouseEvent('click', { bubbles: true }))

  it('a click starts editing it with all of its text selected, so typing replaces it', () => {
    const v = fieldView('USt-IdNr.')
    clickField(v)
    expect(v.dom.querySelector('.doc-protected')!.classList.contains('doc-content-editing')).toBe(true)
    expect(document.getSelection()!.toString()).toBe('USt-IdNr.')
  })

  it('a click on a field holding other text leaves it as the editor does', () => {
    const v = fieldView('DE123456789')
    clickField(v)
    expect(v.dom.querySelector('.doc-protected')!.classList.contains('doc-content-editing')).toBe(false)
  })
})
