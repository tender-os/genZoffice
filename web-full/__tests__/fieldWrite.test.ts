// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { history, undo } from '@tiptap/pm/history'
import { EditorState, NodeSelection, type Transaction } from '@tiptap/pm/state'
import { EditorView } from '@tiptap/pm/view'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { BlockRef, FieldGap } from '../embed-protocol'
import { applyFieldValue, HOST_WRITE_META, placeCaret, type FieldValueResult } from '../fieldWrite'
import { fieldNodeView } from './fieldNodeView'
import { cell, doc, fieldPara, para, placeholderPara, schema, table } from './markerSchema'

let view: EditorView | null = null
afterEach(() => { view?.dom.remove(); view?.destroy(); view = null })

function viewOf(...blocks: PmNode[]) {
  const mount = document.createElement('div')
  document.body.append(mount)
  view = new EditorView(mount, { state: EditorState.create({ schema, doc: doc(...blocks) }), nodeViews: { docProtected: fieldNodeView } })
  return view
}
const paragraphBlock = (paragraph: number): BlockRef => ({ paragraph, scope: 'paragraph' })
const cellBlock = (paragraph: number): BlockRef => ({ paragraph, scope: 'cell' })
// the text-based assertions predate gaps, which have their own tests below
const withoutGap = (result: FieldValueResult) => (result.ok ? { ok: true } : result)
const write = (v: EditorView, block: BlockRef, previous: string, value: string) => withoutGap(applyFieldValue(v, { block, previous, value }))
const texts = (v: EditorView) => {
  const out: string[] = []
  v.state.doc.descendants((node) => {
    if (node.isTextblock) out.push(node.textContent)
    if (node.attrs.fieldDisplay) out.push(`[${node.attrs.fieldDisplay.left}]`)
    return !node.isTextblock
  })
  return out
}
const arial = schema.marks.docTextStyle.create({ font: 'Arial' })
const bold = schema.marks.docTextStyle.create({ bold: true })

describe('applyFieldValue', () => {
  it('replaces the single occurrence in a paragraph', () => {
    const v = viewOf(para(0, 'Name:'), para(1, 'Firma: Muster GmbH'))
    expect(write(v, paragraphBlock(1), 'Muster GmbH', 'Beispiel AG')).toEqual({ ok: true })
    expect(texts(v)).toEqual(['Name:', 'Firma: Beispiel AG'])
  })

  it('maps text offsets over runs and inline leaves, keeping the formatting at the replaced text', () => {
    const runs = [schema.text('Ort: ', [bold]), schema.text('Ber', [arial]), schema.text('lin', [arial]), schema.nodes.docBreak.create(), schema.text('Mitte')]
    const v = viewOf(schema.nodes.docParagraph.create({ docxPara: 0 }, runs))
    expect(write(v, paragraphBlock(0), 'Berlin', 'Köln')).toEqual({ ok: true })
    const paragraph = v.state.doc.child(0)
    expect(paragraph.textContent).toBe('Ort: KölnMitte')
    expect(paragraph.childCount).toBe(4)
    expect(paragraph.child(1).marks).toEqual([arial])
    expect(paragraph.child(2).type.name).toBe('docBreak')
  })

  it("replaces inside a field passthrough paragraph's left text and keeps other attrs", () => {
    const v = viewOf(para(0, 'Name:'), fieldPara(1, 'Muster Ingenieure GmbH', { runs: [{ text: 'Muster Ingenieure GmbH' }] }))
    expect(write(v, paragraphBlock(1), 'Muster Ingenieure GmbH', '')).toEqual({ ok: true })
    const field = v.state.doc.child(1)
    expect(field.attrs.docxPara).toBe(1)
    expect(field.attrs.styleId).toBe('Formular')
    // the formatted runs describe the old text; the editor's own field commit drops them the same way. The
    // emptied result shows as the field's empty line, so the field keeps an element to type into
    expect(field.attrs.fieldDisplay).toEqual({ kind: 'text', left: '', fontFamily: 'Arial', emptyLine: true })
    expect(v.nodeDOM(v.state.doc.child(0).nodeSize)?.textContent).toBe('\u00a0')

    // restoring the value: the empty result matches an empty previous, and the empty line goes again
    expect(write(v, paragraphBlock(1), '', 'Muster Ingenieure GmbH')).toEqual({ ok: true })
    expect(v.state.doc.child(1).attrs.fieldDisplay).toEqual({ kind: 'text', left: 'Muster Ingenieure GmbH', fontFamily: 'Arial' })
  })

  it("matches an empty legacy field's whitespace result in both shapes the editor shows it", () => {
    const enSpaces = '\u2002'.repeat(5)
    // a paragraph keeps the en spaces of the field result as text
    const v = viewOf(para(0, enSpaces), fieldPara(1, '', { emptyLine: true }), fieldPara(2, 'Muster GmbH'))
    expect(write(v, paragraphBlock(0), enSpaces, '089 123')).toEqual({ ok: true })
    expect(v.state.doc.child(0).textContent).toBe('089 123')
    // a field passthrough shows its result trimmed: whitespace only is an empty result
    expect(write(v, paragraphBlock(1), enSpaces, 'Beispiel AG')).toEqual({ ok: true })
    expect(v.state.doc.child(1).attrs.fieldDisplay.left).toBe('Beispiel AG')
    expect(v.state.doc.child(1).attrs.fieldDisplay.emptyLine).toBeUndefined()
    // and the whitespace around a value is gone from it
    expect(write(v, paragraphBlock(2), '\u00a0Muster GmbH ', 'X')).toEqual({ ok: true })
    expect(v.state.doc.child(2).attrs.fieldDisplay.left).toBe('X')
    // a whitespace-only previous in a paragraph without it stays absent
    const plain = viewOf(para(0, 'Name'))
    expect(write(plain, paragraphBlock(0), enSpaces, 'X')).toEqual({ ok: false, reason: 'absent' })
  })

  it('refuses when previous occurs twice', () => {
    const v = viewOf(para(0, 'ja / ja'))
    expect(write(v, paragraphBlock(0), 'ja', 'nein')).toEqual({ ok: false, reason: 'ambiguous' })
    expect(texts(v)).toEqual(['ja / ja'])
  })

  it('counts overlapping occurrences as several', () => {
    const v = viewOf(para(0, 'aaa'))
    expect(write(v, paragraphBlock(0), 'aa', 'b')).toEqual({ ok: false, reason: 'ambiguous' })
  })

  it('refuses when previous is absent', () => {
    const v = viewOf(para(0, 'Name: Muster'))
    expect(write(v, paragraphBlock(0), 'Beispiel', 'X')).toEqual({ ok: false, reason: 'absent' })
  })

  it('inserts into an empty cell when previous is empty', () => {
    const v = viewOf(table(cell(para(0, 'USt-IdNr.:')), cell(para(1), para(2))))
    expect(write(v, cellBlock(2), '', 'DE123456789')).toEqual({ ok: true })
    expect(texts(v)).toEqual(['USt-IdNr.:', 'DE123456789', ''])
  })

  it('replaces text in the paragraph of a cell that holds it', () => {
    const v = viewOf(table(cell(para(0, 'Telefon')), cell(para(1, 'Zentrale'), para(2, '089 123'))))
    expect(write(v, cellBlock(1), '089 123', '089 456')).toEqual({ ok: true })
    expect(texts(v)).toEqual(['Telefon', 'Zentrale', '089 456'])
  })

  it('refuses a match that spans two paragraphs of a cell', () => {
    const v = viewOf(table(cell(para(0, 'Muster'), para(1, 'GmbH'))))
    expect(write(v, cellBlock(0), 'Muster\nGmbH', 'X')).toEqual({ ok: false, reason: 'notEditable' })
  })

  it('refuses an empty previous in a non-empty paragraph', () => {
    const v = viewOf(para(0, 'Name:'), para(1))
    expect(write(v, paragraphBlock(0), '', 'X')).toEqual({ ok: false, reason: 'absent' })
    expect(v.state.doc.child(0).textContent).toBe('Name:')
  })

  it('writes into an empty paragraph, so emptying it can be undone', () => {
    const v = viewOf(para(0, 'Name:'), para(1))
    expect(write(v, paragraphBlock(1), '', 'X')).toEqual({ ok: true })
    expect(v.state.doc.child(1).textContent).toBe('X')
  })

  it("replaces a block's one form field when asked to, however out of date previous is", () => {
    const v = viewOf(fieldPara(0, 'typed since'), para(2, 'x'), para(3))
    const stale = { block: paragraphBlock(0), previous: 'Muster GmbH', value: 'Neu' }
    expect(withoutGap(applyFieldValue(v, stale))).toEqual({ ok: false, reason: 'absent' })
    expect(withoutGap(applyFieldValue(v, { ...stale, replaces: 'wholeField' }))).toEqual({ ok: true })
    expect(v.state.doc.child(0).attrs.fieldDisplay.left).toBe('Neu')
    // a form field that fills its paragraph loads as the paragraph's text
    expect(withoutGap(applyFieldValue(v, { block: paragraphBlock(2), previous: 'gone', value: 'Neu', replaces: 'wholeField' }))).toEqual({ ok: true })
    expect(texts(v)).toContain('Neu')
    // and saved empty, as an empty paragraph
    expect(withoutGap(applyFieldValue(v, { block: paragraphBlock(3), previous: '\u2002', value: 'Y', replaces: 'wholeField' }))).toEqual({ ok: true })
    expect(v.state.doc.lastChild?.textContent).toBe('Y')
  })

  it('takes the whole field when asked to, also when previous still matches part of it', () => {
    const v = viewOf(para(0, 'Muster GmbHQ'), para(1, 'x'))
    const block = paragraphBlock(0)
    expect(withoutGap(applyFieldValue(v, { block, previous: 'Muster GmbH', value: 'Neu', replaces: 'wholeField' }))).toEqual({ ok: true })
    expect(v.state.doc.child(0).textContent).toBe('Neu')
    // a value equal to previous still clears what was typed
    const typed = viewOf(para(0, 'Muster GmbHQ'), para(1, 'x'))
    applyFieldValue(typed, { block, previous: 'Muster GmbH', value: 'Muster GmbH', replaces: 'wholeField' })
    expect(typed.state.doc.child(0).textContent).toBe('Muster GmbH')
  })

  it('refuses an index naming two nodes', () => {
    const v = viewOf(para(3, 'Muster'), para(4, 'und'), para(3, 'GmbH'))
    expect(write(v, paragraphBlock(3), 'Muster', 'X')).toEqual({ ok: false, reason: 'noBlock' })
    expect(write(v, paragraphBlock(9), 'Muster', 'X')).toEqual({ ok: false, reason: 'noBlock' })
    // a cell scope outside any table
    const plain = viewOf(para(0, 'Muster'))
    expect(write(plain, cellBlock(0), 'Muster', 'X')).toEqual({ ok: false, reason: 'noBlock' })
  })

  it('refuses a lone field end, which has no result text for the save to write into', () => {
    const v = viewOf(para(0, 'Firma:'), fieldPara(1, '', { emptyLine: true, fieldEndOnly: true }))
    expect(write(v, paragraphBlock(1), '', 'Muster GmbH')).toEqual({ ok: false, reason: 'notEditable' })
    expect(write(v, paragraphBlock(1), '\u2002\u2002', 'Muster GmbH')).toEqual({ ok: false, reason: 'notEditable' })
    expect(v.state.doc.child(1).attrs.fieldDisplay.left).toBe('')
    expect(placeCaret(v, paragraphBlock(1))).toBe(false)
  })

  it("writes Word's empty result (five en spaces) and finds it again for the next write", () => {
    const enSpaces = '\u2002'.repeat(5)
    const v = viewOf(table(cell(para(0, 'Fax:')), cell(para(1, '089 123'))), fieldPara(2, 'Muster GmbH'))
    // reject: the field result becomes Word's empty result
    expect(write(v, cellBlock(1), '089 123', enSpaces)).toEqual({ ok: true })
    expect(write(v, paragraphBlock(2), 'Muster GmbH', enSpaces)).toEqual({ ok: true })
    expect(texts(v)).toEqual(['Fax:', enSpaces, `[${enSpaces}]`])
    expect(v.state.doc.child(1).attrs.fieldDisplay.emptyLine).toBeUndefined()
    // restore: the en spaces are the previous text
    expect(write(v, cellBlock(1), enSpaces, '089 123')).toEqual({ ok: true })
    expect(write(v, paragraphBlock(2), enSpaces, 'Muster GmbH')).toEqual({ ok: true })
    expect(texts(v)).toEqual(['Fax:', '089 123', '[Muster GmbH]'])
  })

  it('refuses a block without editable text', () => {
    const v = viewOf(schema.nodes.docProtected.create({ docxPara: 0, fieldDisplay: { kind: 'tocLine', left: 'Kapitel', right: '3' } }))
    expect(write(v, paragraphBlock(0), 'Kapitel', 'X')).toEqual({ ok: false, reason: 'notEditable' })
  })

  it('marks the transaction as a host write', () => {
    const v = viewOf(para(0, 'Name: Muster'))
    const seen: Transaction[] = []
    const dispatch = vi.fn((tr: Transaction) => { seen.push(tr); v.updateState(v.state.apply(tr)) })
    v.setProps({ dispatchTransaction: dispatch })
    write(v, paragraphBlock(0), 'Muster', 'Beispiel')
    expect(seen).toHaveLength(1)
    expect(seen[0].getMeta(HOST_WRITE_META)).toBe(true)
  })

  it('a two-field paragraph keeps the other field\'s text', () => {
    const v = viewOf(para(0, 'Ort: Berlin Datum: 01.10.2026'), fieldPara(1, 'Ort: München Datum: 02.10.2026'))
    expect(write(v, paragraphBlock(0), 'Berlin', '')).toEqual({ ok: true })
    expect(write(v, paragraphBlock(1), '02.10.2026', '03.10.2026')).toEqual({ ok: true })
    expect(texts(v)).toEqual(['Ort:  Datum: 01.10.2026', '[Ort: München Datum: 03.10.2026]'])
  })

  it('deleting a whole paragraph text keeps the paragraph', () => {
    const v = viewOf(para(0, 'Muster'), para(1, 'Rest'))
    expect(write(v, paragraphBlock(0), 'Muster', '')).toEqual({ ok: true })
    expect(v.state.doc.childCount).toBe(2)
    expect(v.state.doc.child(0).attrs.docxPara).toBe(0)
    expect(texts(v)).toEqual(['', 'Rest'])
  })
})

describe('gap writes', () => {
  const gapWrite = (v: EditorView, block: BlockRef, gap: FieldGap, value: string) =>
    applyFieldValue(v, { block, value, replaces: 'gap', gap })
  const reject = (v: EditorView, block: BlockRef, previous: string) => applyFieldValue(v, { block, previous, value: '' })

  it('answers an emptying write with the shortest text around the emptied place that is unique, and writes it back there', () => {
    const v = viewOf(para(0, 'beim Amtsgericht Leipzig Mitte.'))
    const emptied = reject(v, paragraphBlock(0), 'Leipzig Mitte')
    expect(emptied).toEqual({ ok: true, gap: { before: 't ', after: '.' } })
    expect(texts(v)).toEqual(['beim Amtsgericht .'])
    expect(gapWrite(v, paragraphBlock(0), { before: 't ', after: '.' }, 'Leipzig Mitte')).toEqual({ ok: true, gap: { before: 't ', after: '.' } })
    expect(texts(v)).toEqual(['beim Amtsgericht Leipzig Mitte.'])
  })

  it('anchors a field at the start or the end of its paragraph to that edge', () => {
    const v = viewOf(para(0, 'Muster GmbH, vertreten durch'), para(1, 'Ort: Berlin'))
    expect(reject(v, paragraphBlock(0), 'Muster GmbH')).toEqual({ ok: true, gap: { before: '', after: ',' } })
    expect(reject(v, paragraphBlock(1), 'Berlin')).toEqual({ ok: true, gap: { before: ' ', after: '' } })
    expect(gapWrite(v, paragraphBlock(0), { before: '', after: ',' }, 'Beispiel AG').ok).toBe(true)
    expect(gapWrite(v, paragraphBlock(1), { before: ' ', after: '' }, 'Köln').ok).toBe(true)
    expect(texts(v)).toEqual(['Beispiel AG, vertreten durch', 'Ort: Köln'])
  })

  it('keeps two emptied fields of one paragraph apart', () => {
    const v = viewOf(para(0, 'Ort: Berlin, Datum: 01.10.2026'))
    const place = reject(v, paragraphBlock(0), 'Berlin')
    const date = reject(v, paragraphBlock(0), '01.10.2026')
    expect(texts(v)).toEqual(['Ort: , Datum: '])
    expect(place).toEqual({ ok: true, gap: { before: ' ', after: ',' } })
    expect(date).toEqual({ ok: true, gap: { before: 'm: ', after: '' } })
    if (!place.ok || !date.ok || !place.gap || !date.gap) throw new Error('no gap')
    expect(gapWrite(v, paragraphBlock(0), date.gap, '02.10.2026').ok).toBe(true)
    expect(gapWrite(v, paragraphBlock(0), place.gap, 'Köln').ok).toBe(true)
    expect(texts(v)).toEqual(['Ort: Köln, Datum: 02.10.2026'])
  })

  it('grows the anchor past a word that repeats around the field', () => {
    const v = viewOf(para(0, 'Name Name Muster Name Name'))
    const emptied = reject(v, paragraphBlock(0), 'Muster')
    if (!emptied.ok || !emptied.gap) throw new Error('no gap')
    expect(emptied.gap).toEqual({ before: ' Name ', after: ' Name ' })
    expect(gapWrite(v, paragraphBlock(0), emptied.gap, 'Beispiel').ok).toBe(true)
    expect(texts(v)).toEqual(['Name Name Beispiel Name Name'])
  })

  it('writes beside a label in the same table cell', () => {
    const v = viewOf(table(cell(para(0, 'Sitz')), cell(para(1, 'Sitz: Berlin'), para(2, 'Deutschland'))))
    const emptied = reject(v, cellBlock(1), 'Berlin')
    expect(emptied).toEqual({ ok: true, gap: { before: ' ', after: '\n' } })
    expect(gapWrite(v, cellBlock(1), { before: ' ', after: '\n' }, 'Leipzig').ok).toBe(true)
    expect(texts(v)).toEqual(['Sitz', 'Sitz: Leipzig', 'Deutschland'])
  })

  it('writes the same gap twice in a row, the second replacing the first', () => {
    const v = viewOf(para(0, 'beim Amtsgericht .'))
    const gap = { before: 't ', after: '.' }
    expect(gapWrite(v, paragraphBlock(0), gap, 'Leipzig').ok).toBe(true)
    expect(gapWrite(v, paragraphBlock(0), gap, 'Leipzig Mitte').ok).toBe(true)
    expect(texts(v)).toEqual(['beim Amtsgericht Leipzig Mitte.'])
  })

  it('takes the formatting of the run at that place', () => {
    const runs = [schema.text('beim ', [bold]), schema.text('Amtsgericht ', [arial]), schema.text('.')]
    const v = viewOf(schema.nodes.docParagraph.create({ docxPara: 0 }, runs))
    expect(gapWrite(v, paragraphBlock(0), { before: 't ', after: '.' }, 'Leipzig').ok).toBe(true)
    const paragraph = v.state.doc.child(0)
    expect(paragraph.textContent).toBe('beim Amtsgericht Leipzig.')
    expect(paragraph.child(1).text).toBe('Amtsgericht Leipzig')
    expect(paragraph.child(1).marks).toEqual([arial])
  })

  it('writes into a field passthrough the same way', () => {
    const v = viewOf(fieldPara(0, 'Ort: Berlin'))
    expect(reject(v, paragraphBlock(0), 'Berlin')).toEqual({ ok: true, gap: { before: ' ', after: '' } })
    expect(gapWrite(v, paragraphBlock(0), { before: ' ', after: '' }, 'Köln').ok).toBe(true)
    expect(texts(v)).toEqual(['[Ort: Köln]'])
  })

  it('refuses when the reviewer changed the text around the gap, never guessing', () => {
    const v = viewOf(para(0, 'beim Landgericht .'))
    expect(gapWrite(v, paragraphBlock(0), { before: 't ', after: ':' }, 'X')).toEqual({ ok: false, reason: 'absent' })
    const twice = viewOf(para(0, 'beim Amtsgericht . und Amtsgericht .'))
    expect(gapWrite(twice, paragraphBlock(0), { before: 't ', after: '.' }, 'X')).toEqual({ ok: false, reason: 'ambiguous' })
    expect(texts(twice)).toEqual(['beim Amtsgericht . und Amtsgericht .'])
  })

  it('refuses a gap across two paragraphs of a cell and a block it cannot find', () => {
    const v = viewOf(table(cell(para(0, 'Muster'), para(1, 'GmbH'))))
    expect(gapWrite(v, cellBlock(0), { before: 'Mus', after: 'bH' }, 'X')).toEqual({ ok: false, reason: 'notEditable' })
    expect(gapWrite(v, paragraphBlock(9), { before: '', after: '' }, 'X')).toEqual({ ok: false, reason: 'noBlock' })
  })

  it('answers without a gap when no anchor of up to 32 characters a side is unique', () => {
    const v = viewOf(para(0, `${'x'.repeat(50)}Y${'x'.repeat(50)}`))
    expect(reject(v, paragraphBlock(0), 'Y')).toEqual({ ok: true })
  })

  it('is one transaction, one undo step', () => {
    const state = EditorState.create({ schema, doc: doc(para(0, 'beim Amtsgericht .')), plugins: [history()] })
    const mount = document.createElement('div')
    document.body.append(mount)
    view = new EditorView(mount, { state })
    expect(gapWrite(view, paragraphBlock(0), { before: 't ', after: '.' }, 'Leipzig Mitte').ok).toBe(true)
    undo(view.state, view.dispatch)
    expect(texts(view)).toEqual(['beim Amtsgericht .'])
  })
})

describe('placeCaret', () => {
  const scrolled = vi.fn()
  beforeEach(() => { Element.prototype.scrollIntoView = scrolled })

  it("puts the caret at the end of a paragraph and ends in a cell's last paragraph", () => {
    const v = viewOf(para(0, 'Name: Muster'), fieldPara(1, 'Firma'), table(cell(para(2, 'a'), para(3, 'bc'))))
    v.focus = vi.fn()

    expect(placeCaret(v, paragraphBlock(0))).toBe(true)
    expect(scrolled).toHaveBeenCalledWith({ block: 'nearest', behavior: 'smooth' })
    expect(v.state.selection.empty).toBe(true)
    expect(v.state.selection.from).toBe(1 + 'Name: Muster'.length)
    expect(v.focus).toHaveBeenCalled()

    expect(placeCaret(v, cellBlock(2))).toBe(true)
    const $caret = v.state.selection.$from
    expect($caret.parent.attrs.docxPara).toBe(3)
    expect($caret.parentOffset).toBe(2)

    expect(placeCaret(v, paragraphBlock(7))).toBe(false)
  })

  it('enters a field passthrough the way a double click does, with no NodeSelection on it', () => {
    const v = viewOf(para(0, 'Firma:'), fieldPara(1, 'Muster GmbH'), para(2, 'Ort:'))
    // a click on the field selects it as an object, as in the editor
    v.dispatch(v.state.tr.setSelection(NodeSelection.create(v.state.doc, v.state.doc.child(0).nodeSize)))

    expect(placeCaret(v, paragraphBlock(1))).toBe(true)
    const field = v.nodeDOM(v.state.doc.child(0).nodeSize)
    const text = v.dom.querySelector('.doc-field-text')
    expect(field instanceof HTMLElement && field.classList.contains('doc-content-editing')).toBe(true)
    expect(text?.getAttribute('contenteditable')).toBe('true')
    // the DOM caret at the end of the result
    const range = document.getSelection()!.getRangeAt(0)
    expect(range.collapsed).toBe(true)
    expect(range.startContainer === text || text?.contains(range.startContainer)).toBe(true)
    expect(range.startOffset).toBe(range.startContainer === text ? text.childNodes.length : 'Muster GmbH'.length)
    // ProseMirror's selection moved next to the field: the next key cannot replace the block
    expect(v.state.selection).not.toBeInstanceOf(NodeSelection)
    expect(v.state.selection.$from.parent.attrs.docxPara).toBe(2)
  })

  it("selects a content control's placeholder text at the end of the block so typing replaces it", () => {
    const v = viewOf(para(0, 'Rechtsform'), table(cell(placeholderPara(1, '', 'Rechtsform eingeben'))))
    expect(placeCaret(v, cellBlock(1))).toBe(true)
    const { from, to } = v.state.selection
    expect(v.state.doc.textBetween(from, to)).toBe('Rechtsform eingeben')
  })

  it("selects a form field's default text so typing replaces it", () => {
    const v = viewOf(para(0, 'Ort:'), fieldPara(1, 'Ort eingeben', { defaultText: 'Ort eingeben' }))
    expect(placeCaret(v, paragraphBlock(1))).toBe(true)
    expect(document.getSelection()!.toString()).toBe('Ort eingeben')
  })

  it("selects an empty field's placeholder so typing replaces it", () => {
    const v = viewOf(para(0, 'Firma:'), fieldPara(1, '', { emptyLine: true }))
    expect(placeCaret(v, paragraphBlock(1))).toBe(true)
    expect(document.getSelection()!.toString()).toBe('\u00a0')
  })

  it("selects a whitespace-only result (Word's empty field: five en spaces) so typing replaces it", () => {
    const enSpaces = '\u2002'.repeat(5)
    // in a table cell the field result is the paragraph's text
    const v = viewOf(table(cell(para(0, 'Fax:')), cell(para(1, enSpaces))), fieldPara(2, enSpaces))
    v.focus = vi.fn()
    expect(placeCaret(v, cellBlock(1))).toBe(true)
    expect(v.state.selection.empty).toBe(false)
    expect(v.state.doc.textBetween(v.state.selection.from, v.state.selection.to)).toBe(enSpaces)
    v.dispatch(v.state.tr.insertText('0171 999'))
    expect(texts(v)).toEqual(['Fax:', '0171 999', `[${enSpaces}]`])
    // a body field passthrough whose result is en spaces (written by the host as an empty result)
    expect(placeCaret(v, paragraphBlock(2))).toBe(true)
    expect(document.getSelection()!.toString()).toBe(enSpaces)
  })
})
