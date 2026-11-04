import { describe, expect, it } from 'vitest'
import { EditorState } from '@tiptap/pm/state'
import { toggleBoxAt } from '../boxTick'
import { doc, para, schema } from './markerSchema'

const stateOf = (text: string) => EditorState.create({ schema, doc: doc(para(0, text)) })
const textAfter = (text: string, pos: number) => toggleBoxAt(stateOf(text), pos)?.doc.textContent

describe('toggleBoxAt', () => {
  it('ticks an empty box and clears a ticked one', () => {
    expect(textAfter('☐ Neubau', 1)).toBe('☒ Neubau')
    expect(textAfter('☒ Neubau', 1)).toBe('☐ Neubau')
    expect(textAfter('☑ Neubau', 1)).toBe('☐ Neubau')
  })

  it('reads a click just after the glyph as the glyph', () => {
    expect(textAfter('☐ Neubau', 2)).toBe('☒ Neubau')
  })

  it('reads a click reported just before the line as the box that opens it', () => {
    expect(textAfter('☒ Neubau', 0)).toBe('☐ Neubau')
    expect(textAfter('Neubau', 0)).toBeUndefined()
  })

  it('leaves a click anywhere else to the editor', () => {
    expect(textAfter('☐ Neubau', 5)).toBeUndefined()
    expect(textAfter('Neubau', 1)).toBeUndefined()
  })

  it('keeps a form field on its box, so the save keeps the field', () => {
    const field = schema.marks.instrField.create({ instr: 'FORMCHECKBOX' })
    const paragraph = schema.nodes.docParagraph.create(null, [schema.text('☒', [field]), schema.text(' Neubau')])
    const state = EditorState.create({ schema, doc: doc(paragraph) })

    const box = toggleBoxAt(state, 1)?.doc.nodeAt(1)

    expect(box?.text).toBe('☐')
    expect(box?.marks.map((mark) => mark.type.name)).toEqual(['instrField'])
  })

  it('toggles the box that was clicked when a paragraph has several', () => {
    expect(textAfter('☐ Ja ☐ Nein', 6)).toBe('☐ Ja ☒ Nein')
  })
})
