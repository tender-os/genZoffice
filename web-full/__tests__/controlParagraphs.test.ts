// @vitest-environment jsdom
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx, type ParsedDocFull } from '@genoffice/docx-engine'
import { buildDocx } from '../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan } from '../../apps/docs/src/renderer/editor/convert'
import { editorExtensions } from '../../apps/docs/src/renderer/editor/extensions'
import { editScopePlugin } from '../editScope'
import { placeholderAt } from '../placeholderText'

const STYLE =
  '<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/><w:rPr><w:color w:val="808080"/></w:rPr></w:style>'
const control = (id: string, text: string) =>
  `<w:p><w:sdt><w:sdtPr><w:id w:val="${id}"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>${text}</w:t></w:r></w:sdtContent></w:sdt></w:p>`
const CONTROLS = `${control('1', 'Datum eingeben')}${control('2', 'Ort eintragen')}`
const blockControl = (id: string, text: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="${id}"/><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:sdtContent></w:sdt>`
const BLOCK_CONTROLS = `${blockControl('3', 'Musterstadt')}${blockControl('4', 'Musterweg 1')}`
const IN_CELL = `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${CONTROLS}</w:tc></w:tr></w:tbl>`

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((editor) => editor.destroy()))

async function open(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, extraStylesXml: STYLE }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  editor.registerPlugin(editScopePlugin(vi.fn()), (plugin, plugins) => [plugin, ...plugins])
  editors.push(editor)
  return { parsed, editor }
}

const textPos = (editor: Editor, text: string) => {
  let found = -1
  editor.state.doc.descendants((node, pos) => {
    if (found === -1 && node.isText && node.text?.includes(text))
      found = pos + node.text.indexOf(text)
  })
  return found
}

function selectPlaceholder(editor: Editor, text: string) {
  const range = placeholderAt(editor.state.doc, textPos(editor, text))!
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, range.from, range.to)),
  )
}

/** the saved file's content controls, each as the texts of its paragraphs */
async function savedControlParagraphs(parsed: ParsedDocFull, editor: Editor) {
  const saved = await saveDocx(
    parsed,
    pmDocToSavePlan(editor.getJSON() as never, parsed.blocks).saveBlocks,
  )
  const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
  const controls = new DOMParser()
    .parseFromString(xml, 'application/xml')
    .getElementsByTagName('w:sdt')
  return [...controls].map((control) =>
    [...control.getElementsByTagName('w:p')].map((paragraph) => paragraph.textContent),
  )
}

const savedControls = async (parsed: ParsedDocFull, editor: Editor) =>
  (await savedControlParagraphs(parsed, editor)).length

const controlParagraphTexts = (editor: Editor) => {
  const texts: string[] = []
  editor.state.doc.descendants((node) => {
    if (node.attrs.controlEnter) texts.push(node.textContent)
    return !node.isTextblock
  })
  return texts
}

describe('content controls through a review edit', () => {
  it.each([
    ['a table cell', IN_CELL],
    ['the body', CONTROLS],
  ])(
    'in %s, Backspace at the start of an emptied control keeps its paragraph and both controls',
    async (_, body) => {
      const { parsed, editor } = await open(body)
      selectPlaceholder(editor, 'Datum eingeben')
      editor.view.dispatch(editor.state.tr.insertText('09.11.2026'))
      selectPlaceholder(editor, 'Ort eintragen')
      editor.commands.keyboardShortcut('Backspace')
      editor.commands.keyboardShortcut('Backspace')

      expect(controlParagraphTexts(editor)).toEqual(['09.11.2026', ''])
      expect(await savedControls(parsed, editor)).toBe(2)
    },
  )

  it('Enter in a control paragraph does nothing', async () => {
    const { parsed, editor } = await open(IN_CELL)
    selectPlaceholder(editor, 'Datum eingeben')
    editor.view.dispatch(editor.state.tr.insertText('09.11.2026'))
    editor.commands.enter()

    expect(controlParagraphTexts(editor)).toEqual(['09.11.2026', 'Ort eintragen'])
    expect(await savedControls(parsed, editor)).toBe(2)
  })

  it('typing over a placeholder still works', async () => {
    const { parsed, editor } = await open(IN_CELL)
    selectPlaceholder(editor, 'Ort eintragen')
    editor.view.dispatch(editor.state.tr.insertText('Berlin'))

    expect(controlParagraphTexts(editor)).toEqual(['Datum eingeben', 'Berlin'])
    expect(await savedControls(parsed, editor)).toBe(2)
  })

  it('Enter inside a block-level control does not split it', async () => {
    const { parsed, editor } = await open(BLOCK_CONTROLS)
    editor.commands.setTextSelection(textPos(editor, 'Musterstadt') + 6)
    editor.commands.enter()

    expect(await savedControls(parsed, editor)).toBe(2)
    expect(controlParagraphTexts(editor)).toEqual(['Musterstadt', 'Musterweg 1'])
  })

  it('Backspace at the start of a block-level control does not join it', async () => {
    const { parsed, editor } = await open(BLOCK_CONTROLS)
    editor.commands.setTextSelection(textPos(editor, 'Musterweg 1'))
    editor.commands.keyboardShortcut('Backspace')

    expect(await savedControls(parsed, editor)).toBe(2)
    expect(controlParagraphTexts(editor)).toEqual(['Musterstadt', 'Musterweg 1'])
  })

  it('typing inside a block-level control keeps it', async () => {
    const { parsed, editor } = await open(BLOCK_CONTROLS)
    editor.commands.setTextSelection(textPos(editor, 'Musterweg 1') + 'Musterweg '.length)
    editor.commands.insertContent('2')

    expect(controlParagraphTexts(editor)).toEqual(['Musterstadt', 'Musterweg 21'])
    expect(await savedControls(parsed, editor)).toBe(2)
  })

  describe('a multi-paragraph control (a long answer)', () => {
    const para = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
    const LONG_ANSWER =
      `${para('Vorher')}<w:sdt><w:sdtPr><w:id w:val="9"/><w:richText/></w:sdtPr><w:sdtContent>` +
      `${para('Erste Zeile')}${para('Zweite Zeile')}</w:sdtContent></w:sdt>${para('Nachher')}`

    it('Enter at the end of its last paragraph adds a paragraph inside the one control', async () => {
      const { parsed, editor } = await open(LONG_ANSWER)
      editor.commands.setTextSelection(textPos(editor, 'Zweite Zeile') + 'Zweite Zeile'.length)
      editor.commands.enter()
      editor.commands.insertContent('Dritte Zeile')

      expect(await savedControlParagraphs(parsed, editor)).toEqual([
        ['Erste Zeile', 'Zweite Zeile', 'Dritte Zeile'],
      ])
    })

    it('Enter inside its first paragraph keeps both halves inside', async () => {
      const { parsed, editor } = await open(LONG_ANSWER)
      editor.commands.setTextSelection(textPos(editor, 'Erste Zeile') + 'Erste'.length)
      editor.commands.enter()

      expect(await savedControlParagraphs(parsed, editor)).toEqual([
        ['Erste', ' Zeile', 'Zweite Zeile'],
      ])
    })

    it('Backspace between two of its paragraphs joins them inside the control', async () => {
      const { parsed, editor } = await open(LONG_ANSWER)
      editor.commands.setTextSelection(textPos(editor, 'Zweite Zeile'))
      editor.commands.keyboardShortcut('Backspace')

      expect(await savedControlParagraphs(parsed, editor)).toEqual([['Erste ZeileZweite Zeile']])
    })

    it.each([
      ['its first paragraph', 'Erste Zeile'],
      ['the paragraph after it', 'Nachher'],
    ])('Backspace at the start of %s moves no text across its edge', async (_, text) => {
      const { parsed, editor } = await open(LONG_ANSWER)
      editor.commands.setTextSelection(textPos(editor, text))
      editor.commands.keyboardShortcut('Backspace')

      expect(await savedControlParagraphs(parsed, editor)).toEqual([
        ['Erste Zeile', 'Zweite Zeile'],
      ])
    })
  })

  describe('Enter by control type', () => {
    const para = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
    const PLACEHOLDER = 'Klicken Sie hier, um Text einzugeben.'
    const blockOfLines = (sdtPr: string, lines: string[]) =>
      `<w:sdt><w:sdtPr><w:id w:val="7"/>${sdtPr}</w:sdtPr><w:sdtContent>${lines.map(para).join('')}</w:sdtContent></w:sdt>`
    const blockOf = (sdtPr: string) =>
      `<w:sdt><w:sdtPr><w:id w:val="5"/>${sdtPr}</w:sdtPr><w:sdtContent>${para(PLACEHOLDER)}</w:sdtContent></w:sdt>`
    const inlineOf = (sdtPr: string) =>
      `<w:p><w:sdt><w:sdtPr><w:id w:val="6"/>${sdtPr}</w:sdtPr><w:sdtContent><w:r><w:t>${PLACEHOLDER}</w:t></w:r></w:sdtContent></w:sdt></w:p>`
    /** the saved controls, each as its paragraphs, a line break written as a newline */
    async function savedControlLines(parsed: ParsedDocFull, editor: Editor) {
      const saved = await saveDocx(
        parsed,
        pmDocToSavePlan(editor.getJSON() as never, parsed.blocks).saveBlocks,
      )
      const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
      const controls = new DOMParser()
        .parseFromString(xml, 'application/xml')
        .getElementsByTagName('w:sdt')
      const lineOf = (paragraph: Element) =>
        [...paragraph.getElementsByTagName('*')]
          .flatMap((element) => {
            if (element.tagName === 'w:t') return [element.textContent]
            return element.tagName === 'w:br' ? ['\n'] : []
          })
          .join('')
      return [...controls].map((control) => {
        const paragraphs = [...control.getElementsByTagName('w:p')]
        return paragraphs.length > 0 ? paragraphs.map(lineOf) : [lineOf(control)]
      })
    }
    const inCell = (content: string) =>
      `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${content}</w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${para('Daneben')}</w:tc></w:tr></w:tbl>`
    async function typeLines(bodyXml: string) {
      const { parsed, editor } = await open(`${para('Vorher')}${bodyXml}${para('Nachher')}`)
      editor.commands.setTextSelection(textPos(editor, PLACEHOLDER) + PLACEHOLDER.length)
      editor.commands.keyboardShortcut('Enter')
      editor.commands.insertContent('Zweite Zeile')
      return savedControlLines(parsed, editor)
    }

    it('a one-paragraph rich-text control takes a new paragraph and stays one control', async () => {
      expect(await typeLines(blockOf(''))).toEqual([[PLACEHOLDER, 'Zweite Zeile']])
    })

    it('a plain-text control takes no Enter', async () => {
      expect(await typeLines(blockOf('<w:text/>'))).toEqual([[`${PLACEHOLDER}Zweite Zeile`]])
    })

    it.each([
      ['a block-level', blockOf('<w:text w:multiLine="1"/>')],
      ['an inline', inlineOf('<w:text w:multiLine="1"/>')],
    ])(
      '%s multi-line plain-text control takes a line break in its one paragraph',
      async (_, body) => {
        expect(await typeLines(body)).toEqual([[`${PLACEHOLDER}\nZweite Zeile`]])
      },
    )

    it.each([
      ['in the body', (body: string) => body],
      ['in a table cell', (body: string) => inCell(body)],
    ])('Shift+Enter in a plain-text control %s adds no line break', async (_, place) => {
      const { parsed, editor } = await open(place(blockOf('<w:text/>')))
      editor.commands.setTextSelection(textPos(editor, PLACEHOLDER) + PLACEHOLDER.length)
      editor.commands.keyboardShortcut('Shift-Enter')
      editor.commands.insertContent('Zweite Zeile')

      expect(await savedControlLines(parsed, editor)).toEqual([[`${PLACEHOLDER}Zweite Zeile`]])
    })

    it('an inline rich-text control filling its paragraph takes a line break', async () => {
      expect(await typeLines(inlineOf(''))).toEqual([[`${PLACEHOLDER}\nZweite Zeile`]])
    })

    describe('a block-level control in a table cell', () => {
      const LONG_ANSWER = inCell(
        `${para('Vorher')}${blockOfLines('', ['Erste Zeile', 'Zweite Zeile'])}${para('Nachher')}`,
      )

      it('a rich-text one takes a new paragraph and stays one control', async () => {
        const { parsed, editor } = await open(LONG_ANSWER)
        editor.commands.setTextSelection(textPos(editor, 'Zweite Zeile') + 'Zweite Zeile'.length)
        editor.commands.keyboardShortcut('Enter')
        editor.commands.insertContent('Dritte Zeile')

        expect(await savedControlLines(parsed, editor)).toEqual([
          ['Erste Zeile', 'Zweite Zeile', 'Dritte Zeile'],
        ])
      })

      it.each([
        ['its first paragraph', 'Erste Zeile'],
        ['the paragraph after it', 'Nachher'],
      ])('Backspace at the start of %s moves no text across its edge', async (_, text) => {
        const { parsed, editor } = await open(LONG_ANSWER)
        editor.commands.setTextSelection(textPos(editor, text))
        editor.commands.keyboardShortcut('Backspace')

        expect(await savedControlLines(parsed, editor)).toEqual([['Erste Zeile', 'Zweite Zeile']])
      })

      it('a plain-text one takes no Enter', async () => {
        const { parsed, editor } = await open(inCell(blockOf('<w:text/>')))
        editor.commands.setTextSelection(textPos(editor, PLACEHOLDER) + 'Klicken'.length)
        editor.commands.keyboardShortcut('Enter')

        expect(await savedControlLines(parsed, editor)).toEqual([[PLACEHOLDER]])
      })

      it('Backspace at the start of a second control does not join it into the first', async () => {
        const { parsed, editor } = await open(
          inCell(`${blockOfLines('', ['Erste'])}${blockOfLines('<w:text/>', ['Zweite'])}`),
        )
        editor.commands.setTextSelection(textPos(editor, 'Zweite'))
        editor.commands.keyboardShortcut('Backspace')

        expect(await savedControlLines(parsed, editor)).toEqual([['Erste'], ['Zweite']])
      })
    })
  })
})
