// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import { history, undo } from '@tiptap/pm/history'
import { EditorState } from '@tiptap/pm/state'
import type { Node as PmNode } from '@tiptap/pm/model'
import { rewriteKeepingHistory } from '../../apps/docs/src/renderer/editor/rewrite-keeping-history'
import { doc, para } from './markerSchema'

const open = (live: PmNode) => EditorState.create({ doc: live, plugins: [history()] })
const typeAt = (state: EditorState, pos: number, text: string) => state.apply(state.tr.insertText(text, pos))
const afterSave = (state: EditorState, saved: PmNode) => {
  const tr = state.tr
  expect(rewriteKeepingHistory(tr, saved)).toBe(true)
  return state.apply(tr.setMeta('addToHistory', false))
}
const undone = (state: EditorState) => {
  let after = state
  undo(state, (tr) => { after = state.apply(tr) })
  return after
}

describe('rewriteKeepingHistory', () => {
  it('keeps what the reviewer typed undoable when the save only renumbered paragraphs', () => {
    const typed = typeAt(open(doc(para(0, 'Muster'), para(1, 'GmbH'))), 2, 'X')
    const saved = afterSave(typed, doc(para(5, 'MXuster'), para(6, 'GmbH')))
    expect(saved.doc.child(0).attrs).toEqual(doc(para(5, 'x')).child(0).attrs)
    expect(undone(saved).doc.child(0).textContent).toBe('Muster')
  })

  it('applies text the saved file holds differently, and keeps the undo of other paragraphs', () => {
    const typed = typeAt(open(doc(para(0, 'Muster'), para(1, 'GmbH'))), 2, 'X')
    const saved = afterSave(typed, doc(para(0, 'MXuster'), para(1, 'AG')))
    expect(saved.doc.child(1).textContent).toBe('AG')
    const back = undone(saved)
    expect(back.doc.child(0).textContent).toBe('Muster')
    expect(back.doc.child(1).textContent).toBe('AG')
  })

  it('declines a saved file of another shape and adds no step', () => {
    const state = open(doc(para(0, 'Muster')))
    const tr = state.tr
    expect(rewriteKeepingHistory(tr, doc(para(0, 'Muster'), para(1, 'GmbH')))).toBe(false)
    expect(tr.docChanged).toBe(false)
  })
})
