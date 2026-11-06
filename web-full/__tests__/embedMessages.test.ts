import { describe, expect, it } from 'vitest'
import { acceptsHostMessage, readHostMessage } from '../embedMessages'

describe('acceptsHostMessage', () => {
  it('accepts only the parent on the same origin', () => {
    const parent = {} as Window
    const self = { origin: 'https://app.test', parent }
    expect(acceptsHostMessage({ origin: 'https://app.test', source: parent }, self)).toBe(true)
    expect(acceptsHostMessage({ origin: 'https://evil.test', source: parent }, self)).toBe(false)
    expect(acceptsHostMessage({ origin: 'https://app.test', source: {} as Window }, self)).toBe(false)
    expect(acceptsHostMessage({ origin: 'https://app.test', source: null }, self)).toBe(false)
  })
})

describe('readHostMessage', () => {
  it('reads open and command', () => {
    const bytes = new ArrayBuffer(4)
    expect(readHostMessage({ type: 'vound:open', name: 'a.docx', bytes })).toEqual({ type: 'vound:open', name: 'a.docx', bytes })
    expect(readHostMessage({ type: 'vound:command', command: 'save' })).toEqual({ type: 'vound:command', command: 'save' })
  })

  it('reads vound:open with mode review', () => {
    const bytes = new ArrayBuffer(4)
    expect(readHostMessage({ type: 'vound:open', name: 'a.docx', bytes, mode: 'review' })).toEqual({ type: 'vound:open', name: 'a.docx', bytes, mode: 'review' })
    // any other mode opens in normal mode, as without one
    expect(readHostMessage({ type: 'vound:open', name: 'a.docx', bytes, mode: 'edit' })).toEqual({ type: 'vound:open', name: 'a.docx', bytes })
    expect(readHostMessage({ type: 'vound:open', name: 'a.docx', bytes, mode: null })).toEqual({ type: 'vound:open', name: 'a.docx', bytes })
  })

  it('reads vound:open with mode view', () => {
    const bytes = new ArrayBuffer(4)
    expect(readHostMessage({ type: 'vound:open', name: 'a.docx', bytes, mode: 'view' })).toEqual({ type: 'vound:open', name: 'a.docx', bytes, mode: 'view' })
  })

  it('reads setFieldValue for a paragraph and a cell block', () => {
    const paragraph = { type: 'vound:setFieldValue', requestId: 'r1', block: { paragraph: 4, scope: 'paragraph' }, previous: 'Berlin', value: 'München' }
    expect(readHostMessage(paragraph)).toEqual(paragraph)
    const cellWrite = { type: 'vound:setFieldValue', requestId: 'r2', block: { paragraph: 0, scope: 'cell' }, previous: '', value: '2026' }
    expect(readHostMessage(cellWrite)).toEqual(cellWrite)
    // only the known fields survive
    expect(readHostMessage({ ...cellWrite, extra: 1, block: { paragraph: 0, scope: 'cell', table: 2 } })).toEqual(cellWrite)
  })

  it('reads a gap write, which needs its gap and no previous', () => {
    const gapWrite = { type: 'vound:setFieldValue', requestId: 'r3', block: { paragraph: 2, scope: 'paragraph' }, value: 'Leipzig', replaces: 'gap', gap: { before: 't ', after: '.' } }
    expect(readHostMessage(gapWrite)).toEqual(gapWrite)
    expect(readHostMessage({ ...gapWrite, previous: 'x', gap: { before: '', after: '', extra: 1 } })).toEqual({ ...gapWrite, gap: { before: '', after: '' } })
    expect(readHostMessage({ ...gapWrite, gap: undefined })).toBeNull()
    expect(readHostMessage({ ...gapWrite, gap: { before: 't ' } })).toBeNull()
    expect(readHostMessage({ ...gapWrite, value: null })).toBeNull()
  })

  it('rejects setFieldValue with a negative paragraph or a non-string previous', () => {
    const valid = { type: 'vound:setFieldValue', requestId: 'r1', block: { paragraph: 4, scope: 'paragraph' }, previous: 'a', value: 'b' }
    expect(readHostMessage({ ...valid, block: { paragraph: -1, scope: 'paragraph' } })).toBeNull()
    expect(readHostMessage({ ...valid, block: { paragraph: 1.5, scope: 'paragraph' } })).toBeNull()
    expect(readHostMessage({ ...valid, block: { paragraph: 4, scope: 'row' } })).toBeNull()
    expect(readHostMessage({ ...valid, block: null })).toBeNull()
    expect(readHostMessage({ ...valid, previous: null })).toBeNull()
    expect(readHostMessage({ ...valid, previous: 3 })).toBeNull()
    expect(readHostMessage({ ...valid, value: undefined })).toBeNull()
    expect(readHostMessage({ ...valid, requestId: 7 })).toBeNull()
  })

  it('reads placeCaret', () => {
    expect(readHostMessage({ type: 'vound:placeCaret', block: { paragraph: 2, scope: 'cell' } })).toEqual({ type: 'vound:placeCaret', block: { paragraph: 2, scope: 'cell' } })
    expect(readHostMessage({ type: 'vound:placeCaret', block: { paragraph: '2', scope: 'cell' } })).toBeNull()
    expect(readHostMessage({ type: 'vound:placeCaret' })).toBeNull()
  })

  it('reads markers with label and actions, defaulting both when absent', () => {
    const labelled = { id: 'slot:s1', paragraph: 3, scope: 'cell', tone: 'proposed', label: 'Ort', actions: ['accept', 'reject', 'fill', 'ask', 'restore'] }
    expect(readHostMessage({ type: 'vound:markers', markers: [labelled] })).toEqual({ type: 'vound:markers', markers: [labelled] })
    const bare = { id: 'slot:s2', paragraph: 5, scope: 'paragraph', tone: 'gap' }
    expect(readHostMessage({ type: 'vound:markers', markers: [bare] })).toEqual({ type: 'vound:markers', markers: [{ ...bare, label: 'slot:s2', actions: [] }] })
  })

  it('reads an otherParty marker', () => {
    const marker = { id: 'slot:s1', paragraph: 3, scope: 'paragraph', tone: 'otherParty', label: 'Ort', actions: ['ask'] }
    expect(readHostMessage({ type: 'vound:markers', markers: [marker] })).toEqual({ type: 'vound:markers', markers: [marker] })
  })

  it('rejects a marker with an unknown action', () => {
    const marker = { id: 'slot:s1', paragraph: 3, scope: 'cell', tone: 'proposed', label: 'Ort', actions: ['accept'] }
    expect(readHostMessage({ type: 'vound:markers', markers: [{ ...marker, actions: ['accept', 'delete'] }] })).toBeNull()
    expect(readHostMessage({ type: 'vound:markers', markers: [{ ...marker, actions: 'accept' }] })).toBeNull()
    expect(readHostMessage({ type: 'vound:markers', markers: [{ ...marker, label: 4 }] })).toBeNull()
  })

  it('reads markers, scrollTo and focus, and rejects a malformed marker', () => {
    const marker = { id: 'slot:s1', paragraph: 3, scope: 'cell', tone: 'conditional', label: 'Datum', actions: ['accept'] }
    expect(readHostMessage({ type: 'vound:markers', markers: [marker] })).toEqual({ type: 'vound:markers', markers: [marker] })
    expect(readHostMessage({ type: 'vound:markers', markers: [] })).toEqual({ type: 'vound:markers', markers: [] })
    const withOne = (patch: object) => readHostMessage({ type: 'vound:markers', markers: [marker, { ...marker, ...patch }] })
    expect(withOne({ paragraph: 1.5 })).toBeNull()
    expect(withOne({ paragraph: -1 })).toBeNull()
    expect(withOne({ paragraph: '3' })).toBeNull()
    expect(withOne({ tone: 'pink' })).toBeNull()
    expect(withOne({ scope: 'row' })).toBeNull()
    expect(withOne({ id: 7 })).toBeNull()
    expect(readHostMessage({ type: 'vound:markers', markers: [{ paragraph: 3, scope: 'cell', tone: 'gap' }] })).toBeNull()
    expect(readHostMessage({ type: 'vound:markers', markers: 'none' })).toBeNull()
    expect(readHostMessage({ type: 'vound:markers', markers: [null] })).toBeNull()

    expect(readHostMessage({ type: 'vound:scrollTo', id: 'slot:s2' })).toEqual({ type: 'vound:scrollTo', id: 'slot:s2' })
    expect(readHostMessage({ type: 'vound:scrollTo', id: null })).toBeNull()
    expect(readHostMessage({ type: 'vound:focus', id: 'slot:s2' })).toEqual({ type: 'vound:focus', id: 'slot:s2' })
    expect(readHostMessage({ type: 'vound:focus', id: null })).toEqual({ type: 'vound:focus', id: null })
    expect(readHostMessage({ type: 'vound:focus' })).toBeNull()
  })

  it('keeps only the known marker fields', () => {
    const msg = readHostMessage({ type: 'vound:markers', markers: [{ id: 'a', paragraph: 0, scope: 'paragraph', tone: 'gap', label: 'A', actions: [], extra: 1 }] })
    expect(msg).toEqual({ type: 'vound:markers', markers: [{ id: 'a', paragraph: 0, scope: 'paragraph', tone: 'gap', label: 'A', actions: [] }] })
  })

  it('ignores anything else', () => {
    expect(readHostMessage(null)).toBeNull()
    expect(readHostMessage('vound:open')).toBeNull()
    expect(readHostMessage({ type: 'vound:open', name: 'a.docx', bytes: 'not bytes' })).toBeNull()
    expect(readHostMessage({ type: 'vound:open', bytes: new ArrayBuffer(1) })).toBeNull()
    expect(readHostMessage({ type: 'vound:command', command: 'undo' })).toBeNull()
    expect(readHostMessage({ type: 'vound:command', command: 'save-as' })).toBeNull()
    expect(readHostMessage({ type: 'other' })).toBeNull()
  })
})
