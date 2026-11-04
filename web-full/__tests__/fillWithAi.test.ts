// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import { Schema } from '@tiptap/pm/model'
import { EditorState, TextSelection } from '@tiptap/pm/state'
import { EditorView } from '@tiptap/pm/view'
import { fillWithAiLabel, fillWithAiMessage } from '../fillWithAi'

const schema = new Schema({
  nodes: { doc: { content: 'paragraph+' }, paragraph: { content: 'text*', toDOM: () => ['p', 0] }, text: {} },
})
const viewAt = (from: number, to = from) => {
  const doc = schema.node('doc', null, [
    schema.node('paragraph', null, [schema.text('Firma Muster')]),
    schema.node('paragraph'),
  ])
  const selection = TextSelection.create(doc, from, to)
  return new EditorView(document.createElement('div'), { state: EditorState.create({ doc, selection }) })
}

describe('fillWithAiLabel', () => {
  it('names the entry in the host language, in English otherwise', () => {
    expect(fillWithAiLabel('de')).toBe('Mit KI ausfüllen')
    expect(fillWithAiLabel('ar')).toBe('Fill with AI')
  })
})

describe('fillWithAiMessage', () => {
  it('reports the selection and its paragraph', () => {
    expect(fillWithAiMessage(viewAt(7, 13))).toMatchObject({ selected: 'Muster', paragraph: 'Firma Muster' })
  })

  it('reports the text before an empty line', () => {
    expect(fillWithAiMessage(viewAt(15))).toMatchObject({ selected: '', paragraph: 'Firma Muster' })
  })
})
