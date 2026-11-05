// @vitest-environment jsdom
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type ParsedDocFull } from '@genoffice/docx-engine'
import { buildDocx } from '../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../../apps/docs/src/renderer/editor/convert'
import { editorExtensions } from '../../apps/docs/src/renderer/editor/extensions'
import type { BlockRef, FieldGap } from '../embed-protocol'
import { applyFieldValue } from '../fieldWrite'

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p>${run(text)}</w:p></w:tc>`
const BODY =
  `<w:p>${run('beim Amtsgericht Leipzig Mitte.')}</w:p>` +
  `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell('Sitz')}${cell('Sitz: Berlin')}</w:tr></w:tbl>` +
  `<w:p>${run('Ende')}</w:p>`
const court: BlockRef = { paragraph: 0, scope: 'paragraph' }
const seat: BlockRef = { paragraph: 2, scope: 'cell' }

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((editor) => editor.destroy()))

function open(parsed: ParsedDocFull) {
  const editor = new Editor({ element: document.createElement('div'), extensions: editorExtensions })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  editors.push(editor)
  return editor
}

async function save(parsed: ParsedDocFull, editor: Editor) {
  const plan = pmDocToSavePlan(editor.getJSON() as never, parsed.blocks)
  return saveDocx(parsed, plan.saveBlocks)
}

const documentXml = async (bytes: Uint8Array) => (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')

function gapOf(result: ReturnType<typeof applyFieldValue>): FieldGap {
  if (!result.ok || !result.gap) throw new Error(`no gap: ${JSON.stringify(result)}`)
  return result.gap
}

describe('a gap through save and reload', () => {
  it('writes an emptied field back after the document was saved and opened again, byte for byte as before', async () => {
    const original = await buildDocx({ bodyXml: BODY })
    const first = await parseDocx(original)
    const editor = open(first)
    const courtGap = gapOf(applyFieldValue(editor.view, { block: court, previous: 'Leipzig Mitte', value: '' }))
    const seatGap = gapOf(applyFieldValue(editor.view, { block: seat, previous: 'Berlin', value: '' }))
    const emptied = await save(first, editor)
    expect(await documentXml(emptied)).toContain('beim Amtsgericht .</w:t>')

    const reopened = await parseDocx(emptied)
    const again = open(reopened)
    expect(applyFieldValue(again.view, { block: court, value: 'Leipzig Mitte', replaces: 'gap', gap: courtGap }).ok).toBe(true)
    expect(applyFieldValue(again.view, { block: seat, value: 'Berlin', replaces: 'gap', gap: seatGap }).ok).toBe(true)
    expect(await documentXml(await save(reopened, again))).toBe(await documentXml(original))
  })
})
