// @vitest-environment jsdom
import { undo } from '@tiptap/pm/history'
import { TextSelection } from '@tiptap/pm/state'
import { describe, expect, it } from 'vitest'
import { applyFieldStyle } from '../fieldStyle'
import { FIELD_TEXT_SELECTOR, HOST_WRITE_META, placeCaret } from '../fieldWrite'
import { CALIBRI_11, cell, city, faceOfText, FIELD_END, fieldOpen, firm, open, reopened, run, runOf, savedXml, selectText, styleOfText, type } from './fieldStyle.testkit'

describe('applyFieldStyle', () => {
  it('restyles the whole field when nothing in it is selected, and saves the run properties', async () => {
    const { parsed, editor } = await open()
    const result = applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: { fontFamily: 'Arial', fontSize: 12 } })
    expect(result).toEqual({ ok: true, style: { fontFamily: 'Arial', fontSize: 12, sizing: 'own' } })
    const xml = await savedXml(parsed, editor)
    expect(runOf(xml, 'Muster GmbH')).toContain('w:ascii="Arial"')
    expect(runOf(xml, 'Muster GmbH')).toContain('w:hAnsi="Arial"')
    expect(runOf(xml, 'Muster GmbH')).toContain('<w:sz w:val="24"/>')
    expect(runOf(xml, 'Muster GmbH')).toContain('<w:szCs w:val="24"/>')
    expect(runOf(xml, 'Firma: ')).not.toContain('Arial')
  })

  it('restyles only the selection when it lies inside the field', async () => {
    const { parsed, editor } = await open()
    selectText(editor, 'GmbH')
    applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: { fontSize: 14 } })
    const xml = await savedXml(parsed, editor)
    expect(runOf(xml, 'GmbH')).toContain('<w:sz w:val="28"/>')
    expect(runOf(xml, 'Muster ')).toContain('<w:sz w:val="22"/>')
  })

  it('keeps what the run had and changes only the property asked for', async () => {
    const { parsed, editor } = await open()
    applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: { fontSize: 9 } })
    const field = runOf(await savedXml(parsed, editor), 'Muster GmbH')
    expect(field).toContain('w:ascii="Calibri"')
    expect(field).toContain('<w:sz w:val="18"/>')
  })

  it('restyles a form text field in a table cell, found by its gap', async () => {
    const { parsed, editor } = await open()
    const result = applyFieldStyle(editor.view, { block: city, replaces: 'gap', gap: { before: '', after: '' }, style: { fontFamily: 'Courier New' } })
    expect(result.ok).toBe(true)
    const xml = await savedXml(parsed, editor)
    expect(runOf(xml, 'Berlin')).toContain('w:ascii="Courier New"')
    expect(xml).toContain('FORMTEXT')
  })

  it('is one undoable host write', async () => {
    const { editor } = await open()
    const metas: unknown[] = []
    editor.on('transaction', ({ transaction }) => { if (transaction.docChanged) metas.push(transaction.getMeta(HOST_WRITE_META)) })
    const before = editor.state.doc
    applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: { fontFamily: 'Arial', fontSize: 12 } })
    expect(metas).toEqual([true])
    undo(editor.state, editor.view.dispatch)
    expect(editor.state.doc.eq(before)).toBe(true)
  })

  it('answers with the style the field shows and changes nothing for an empty style', async () => {
    const { editor } = await open()
    const before = editor.state.doc
    expect(applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: {} })).toEqual({ ok: true, style: { fontFamily: 'Calibri', fontSize: 11, sizing: 'own' } })
    expect(editor.state.doc).toBe(before)
  })

  it('leaves out a property the field shows in more than one way', async () => {
    const { editor } = await open()
    selectText(editor, 'GmbH')
    applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: { fontSize: 14 } })
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1)))
    expect(applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: {} })).toEqual({ ok: true, style: { fontFamily: 'Calibri', sizing: 'own' } })
  })

  it('restyles a content control in a paragraph and in a cell, through save and reload', async () => {
    const control = (text: string) => `<w:sdt><w:sdtPr><w:id w:val="7"/><w:text/></w:sdtPr><w:sdtContent>${run(text, CALIBRI_11)}</w:sdtContent></w:sdt>`
    const body =
      `<w:p>${run('Ort: ')}${control('Berlin')}</w:p>` +
      `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
      `<w:tr>${cell(run('Firma'))}${cell(control('Muster GmbH'))}</w:tr></w:tbl><w:p>${run('Ende')}</w:p>`
    const { parsed, editor } = await open(body)
    expect(applyFieldStyle(editor.view, { block: firm, previous: 'Berlin', style: { fontFamily: 'Arial' } }).ok).toBe(true)
    expect(applyFieldStyle(editor.view, { block: { paragraph: 2, scope: 'cell' }, previous: 'Muster GmbH', style: { fontSize: 9 } }).ok).toBe(true)
    const { xml, again } = await reopened(parsed, editor)
    expect(xml.match(/<w:sdt>/g)).toHaveLength(2)
    expect(runOf(xml, 'Berlin')).toContain('w:ascii="Arial"')
    expect(runOf(xml, 'Muster GmbH')).toContain('<w:sz w:val="18"/>')
    expect(applyFieldStyle(again.editor.view, { block: firm, previous: 'Berlin', style: {} })).toEqual({ ok: true, style: { fontFamily: 'Arial', fontSize: 11, sizing: 'own' } })
    expect(applyFieldStyle(again.editor.view, { block: { paragraph: 2, scope: 'cell' }, previous: 'Muster GmbH', style: {} })).toEqual({ ok: true, style: { fontFamily: 'Calibri', fontSize: 9, sizing: 'own', verticalAlignment: 'top' } })
  })

  it('restyles the result of a form field shown as one block, and saves it there only', async () => {
    const dropDown = '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="D"/><w:ddList><w:listEntry w:val="ja"/></w:ddList></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' + run('ja') + FIELD_END
    const { parsed, editor } = await open(`<w:p>${run('Firma: ', CALIBRI_11)}${fieldOpen}${run('Muster GmbH', CALIBRI_11)}${FIELD_END}${run(' ')}${dropDown}</w:p>`)
    expect(editor.state.doc.child(0).attrs.fieldDisplay).not.toBeNull()
    const result = applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: { fontFamily: 'Arial', fontSize: 12 } })
    expect(result).toEqual({ ok: true, style: { fontFamily: 'Arial', fontSize: 12, sizing: 'own' } })
    const shown = editor.view.nodeDOM(0) as HTMLElement
    expect(shown.textContent).toContain('Muster GmbH')
    const { xml, again } = await reopened(parsed, editor)
    expect(runOf(xml, 'Muster GmbH')).toContain('w:ascii="Arial"')
    expect(runOf(xml, 'Muster GmbH')).toContain('<w:sz w:val="24"/>')
    expect(runOf(xml, 'Firma: ')).toContain('w:ascii="Calibri"')
    expect(xml).toContain('FORMDROPDOWN')
    expect(applyFieldStyle(again.editor.view, { block: firm, previous: 'Muster GmbH', style: {} })).toEqual({ ok: true, style: { fontFamily: 'Arial', fontSize: 12, sizing: 'own' } })
  })

  it('keeps the face set on part of a form field block when the reviewer types into it', async () => {
    const dropDown = '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="D"/><w:ddList><w:listEntry w:val="ja"/></w:ddList></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' + run('ja') + FIELD_END
    const { parsed, editor } = await open(`<w:p>${run('Firma: ', CALIBRI_11)}${fieldOpen}${run('Muster GmbH', CALIBRI_11)}${FIELD_END}${run(' ')}${dropDown}</w:p>`)
    applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: { fontFamily: 'Arial' } })
    expect(placeCaret(editor.view, firm)).toBe(true)
    const text = (editor.view.nodeDOM(0) as HTMLElement).querySelector(FIELD_TEXT_SELECTOR) as HTMLElement
    const result = [...text.childNodes].find((child) => child.textContent === 'Muster GmbH')!
    result.textContent = 'Muster GmbH & Co.'
    text.dispatchEvent(new Event('input', { bubbles: true }))
    window.dispatchEvent(new Event('ai-docs-commit-tables'))
    const display = editor.state.doc.child(0).attrs.fieldDisplay
    expect(display.left).toContain('Muster GmbH & Co.')
    expect(display.runs.find((r: { text: string }) => r.text.includes('Muster'))).toMatchObject({ text: 'Muster GmbH & Co.', fontAscii: 'Arial' })
    expect(display.runs.find((r: { text: string }) => r.text.includes('Firma'))?.fontAscii).toBe('Calibri')
    const { xml } = await reopened(parsed, editor)
    expect(runOf(xml, 'Muster GmbH &amp; Co.')).toContain('w:ascii="Arial"')
    expect(runOf(xml, 'Firma: ')).toContain('w:ascii="Calibri"')
  })

  it('restyles an empty form field the field index names, and the text typed into it takes the style', async () => {
    const blank = `${fieldOpen}${run('\u2002'.repeat(5))}${FIELD_END}`
    const { parsed, editor } = await open(`<w:p>${run('Ort: ')}${blank}${run(' Datum: ')}${blank}</w:p>`)
    expect(applyFieldStyle(editor.view, { block: firm, previous: '\u2002'.repeat(5), formText: 1, style: { fontFamily: 'Arial' } }).ok).toBe(true)
    let end = 0
    editor.state.doc.descendants((node, pos) => { if (node.isText) end = pos + node.nodeSize })
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, end)))
    type(editor, 'Mai')
    expect(faceOfText(editor, 'Mai')).toBe('Arial')
    const { xml } = await reopened(parsed, editor)
    expect(runOf(xml, 'Mai')).toContain('w:ascii="Arial"')
  })

  it.each([
    ['its index', { formText: 0 }],
    ['being the one empty field of its line', {}],
  ])('restyles a form field saved empty, found by %s, and the text typed into it takes the style', async (_, hint) => {
    const { parsed, editor } = await open(`<w:p>${run('Ort: ')}${fieldOpen}${run('')}${FIELD_END}</w:p>`)
    expect(applyFieldStyle(editor.view, { block: firm, previous: '', ...hint, style: { fontSize: 14 } }).ok).toBe(true)
    type(editor, 'Köln')
    expect(styleOfText(editor, 'Köln')?.sizeHalfPoints).toBe(28)
    const { xml } = await reopened(parsed, editor)
    expect(runOf(xml, 'Köln')).toContain('<w:sz w:val="28"/>')
    expect(xml).toContain('FORMTEXT')
  })

  it('picks the empty field the caret is in when a line has several', async () => {
    const control = `<w:sdt><w:sdtPr><w:id w:val="3"/><w:text/></w:sdtPr><w:sdtContent>${run('')}</w:sdtContent></w:sdt>`
    const { editor } = await open(`<w:p>${run('Von: ')}${control}${run(' bis: ')}${control}</w:p>`)
    const fillers: number[] = []
    editor.state.doc.descendants((node, pos) => { if (node.isText && node.text === '\u200b') fillers.push(pos) })
    expect(fillers).toHaveLength(2)
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1)))
    expect(applyFieldStyle(editor.view, { block: firm, previous: '', style: { fontSize: 9 } }).ok).toBe(false)
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, fillers[1])))
    expect(applyFieldStyle(editor.view, { block: firm, previous: '', style: { fontSize: 9 } }).ok).toBe(true)
    expect(editor.state.selection.from).toBe(fillers[1])
    expect(editor.state.storedMarks?.find((mark) => mark.type.name === 'docTextStyle')?.attrs.sizeHalfPoints).toBe(18)
  })

  it('stores the style at the end of a label line whose field has no text yet, for the text typed next', async () => {
    const { parsed, editor } = await open(`<w:p>${run('Name: ')}</w:p>`)
    const result = applyFieldStyle(editor.view, { block: firm, previous: '', style: { fontFamily: 'Courier New' } })
    expect(result.ok).toBe(true)
    type(editor, 'Muster')
    expect(faceOfText(editor, 'Muster')).toBe('Courier New')
    const { xml } = await reopened(parsed, editor)
    expect(runOf(xml, 'Muster')).toContain('w:ascii="Courier New"')
  })

  it('refuses a target that names no place, and leaves the document as it was', async () => {
    const { editor } = await open()
    const before = editor.state.doc
    expect(applyFieldStyle(editor.view, { block: firm, previous: 'Beispiel AG', style: { fontSize: 12 } })).toEqual({ ok: false, reason: 'absent' })
    expect(applyFieldStyle(editor.view, { block: { paragraph: 99, scope: 'paragraph' }, previous: 'x', style: { fontSize: 12 } })).toEqual({ ok: false, reason: 'noBlock' })
    expect(editor.state.doc).toBe(before)
  })
})
