// @vitest-environment jsdom
import { undo } from '@tiptap/pm/history'
import { describe, expect, it } from 'vitest'
import { applyFieldStyle } from '../fieldStyle'
import type { BlockRef } from '../embed-protocol'
import { cell, FIELD_END, fieldOpen, firm, open, reopened, run, runOf } from './fieldStyle.testkit'

const city: BlockRef = { paragraph: 1, scope: 'cell' }
const SIZE_14 = '<w:rPr><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
const table = (inner: string) =>
  `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell(run('Ort'))}${cell(inner)}</w:tr></w:tbl><w:p>${run('Ende')}</w:p>`
const formText = (value: string, rPr = '') => `${fieldOpen}${run(value, rPr)}${FIELD_END}`
const dropDown =
  '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="D"/><w:ddList><w:listEntry w:val="ja"/></w:ddList></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  run('ja') +
  FIELD_END

/** the w:p holding the text */
const paragraphOf = (xml: string, text: string) => {
  const at = xml.indexOf(text)
  return xml.slice(Math.max(xml.lastIndexOf('<w:p>', at), xml.lastIndexOf('<w:p ', at)), xml.indexOf('</w:p>', at))
}
const cellOf = (xml: string, text: string) => {
  const at = xml.indexOf(`>${text}</w:t>`)
  return xml.slice(xml.lastIndexOf('<w:tc>', at), xml.indexOf('</w:tc>', at))
}

describe('field alignment and size override', () => {
  it.each([
    ['a plain line', `<w:p>${run('Firma: ')}${run('Muster GmbH')}</w:p>`],
    ['a form field line', `<w:p>${run('Firma: ')}${formText('Muster GmbH')}</w:p>`],
    ['a form field block', `<w:p>${run('Firma: ')}${formText('Muster GmbH')}${run(' ')}${dropDown}</w:p>`],
  ])('aligns the paragraph of a field on %s, through save and reload', async (_, body) => {
    const { parsed, editor } = await open(body)
    expect(applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: { alignment: 'center' } })).toMatchObject({
      ok: true,
      style: { alignment: 'center' },
    })
    const { xml, again } = await reopened(parsed, editor)
    expect(paragraphOf(xml, 'Muster GmbH')).toContain('<w:jc w:val="center"/>')
    expect(xml.includes('FORMTEXT')).toBe(body.includes('FORMTEXT'))
    expect(applyFieldStyle(again.editor.view, { block: firm, previous: 'Muster GmbH', style: {} })).toMatchObject({ style: { alignment: 'center' } })
  })

  it('aligns a form field in a table cell both ways, through save and reload', async () => {
    const { parsed, editor } = await open(table(formText('Berlin')))
    const result = applyFieldStyle(editor.view, { block: city, previous: 'Berlin', style: { alignment: 'right', verticalAlignment: 'bottom' } })
    expect(result).toMatchObject({ ok: true, style: { alignment: 'right', verticalAlignment: 'bottom' } })
    const { xml, again } = await reopened(parsed, editor)
    expect(paragraphOf(xml, 'Berlin')).toContain('<w:jc w:val="right"/>')
    expect(cellOf(xml, 'Berlin')).toContain('<w:vAlign w:val="bottom"/>')
    expect(xml).toContain('FORMTEXT')
    expect(applyFieldStyle(again.editor.view, { block: city, previous: 'Berlin', style: {} })).toMatchObject({
      style: { alignment: 'right', verticalAlignment: 'bottom' },
    })
  })

  it('has no vertical alignment outside a table cell, and changes nothing for one', async () => {
    const { editor } = await open(`<w:p>${run('Firma: ')}${run('Muster GmbH')}</w:p>`)
    const before = editor.state.doc
    const result = applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: { verticalAlignment: 'center' } })
    expect(result.ok && result.style.verticalAlignment).toBeUndefined()
    expect(editor.state.doc.eq(before)).toBe(true)
    const top = await open(table(run('Berlin')))
    expect(applyFieldStyle(top.editor.view, { block: city, previous: 'Berlin', style: {} })).toMatchObject({ style: { verticalAlignment: 'top' } })
  })

  it.each([
    ['a plain line', `<w:p>${run('Firma: ')}${run('Muster GmbH', SIZE_14)}</w:p>`],
    ['a form field line', `<w:p>${run('Firma: ')}${formText('Muster GmbH', SIZE_14)}</w:p>`],
    ['a form field block', `<w:p>${run('Firma: ')}${formText('Muster GmbH', SIZE_14)}${run(' ')}${dropDown}</w:p>`],
  ])('takes the size override of a field on %s back, so it shows the document size again', async (_, body) => {
    const { parsed, editor } = await open(body)
    expect(applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: {} })).toMatchObject({ style: { sizing: 'own', fontSize: 14 } })
    expect(applyFieldStyle(editor.view, { block: firm, previous: 'Muster GmbH', style: { fontSize: null } })).toMatchObject({
      ok: true,
      style: { sizing: 'document' },
    })
    const { xml, again } = await reopened(parsed, editor)
    expect(runOf(xml, 'Muster GmbH')).not.toContain('<w:sz ')
    expect(runOf(xml, 'Muster GmbH')).not.toContain('<w:szCs ')
    expect(applyFieldStyle(again.editor.view, { block: firm, previous: 'Muster GmbH', style: {} })).toMatchObject({ style: { sizing: 'document' } })
  })

  it('undoes a face, size, alignment and cell alignment change in one step', async () => {
    const { editor } = await open(table(formText('Berlin')))
    const before = editor.state.doc
    applyFieldStyle(editor.view, {
      block: city,
      previous: 'Berlin',
      style: { fontFamily: 'Arial', fontSize: 12, alignment: 'center', verticalAlignment: 'center' },
    })
    expect(editor.state.doc.eq(before)).toBe(false)
    undo(editor.state, editor.view.dispatch)
    expect(editor.state.doc.eq(before)).toBe(true)
  })
})
