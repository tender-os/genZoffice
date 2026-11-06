// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { EditorState } from '@tiptap/pm/state'
import { EditorView } from '@tiptap/pm/view'
import type { Node as PmNode } from '@tiptap/pm/model'
import { findPassages, matchText, pageOfNode, REVEAL_CLASS, revealKey, revealPassage, revealPlugin } from '../reveal'
import { createRevealQueue } from '../revealQueue'
import { cell, doc, para, table } from './markerSchema'

let view: EditorView | null = null
afterEach(() => {
  view?.destroy()
  view = null
})

const textBetween = (document: PmNode, range: { from: number; to: number }) => document.textBetween(range.from, range.to, '\n')

function viewOf(...blocks: PmNode[]) {
  Element.prototype.scrollIntoView = vi.fn()
  view = new EditorView(document.createElement('div'), {
    state: EditorState.create({ doc: doc(...blocks), plugins: [revealPlugin(() => {})] }),
  })
  return view
}

const marked = (v: EditorView) => [...v.dom.querySelectorAll(`.${REVEAL_CLASS}`)].map((node) => node.textContent).join('')

describe('matchText', () => {
  it('ignores case, whitespace and hyphens, folds ligatures and the quote style', () => {
    expect(matchText('Der „Bieter“ ﬁndet\n das Ange-\nbot')).toBe(matchText('der "bieter" findet das angebot'))
    expect(matchText('  \n- ')).toBe('')
  })
})

describe('findPassages', () => {
  it('finds the exact text', () => {
    const document = doc(para(0, 'Erklärung zur Eignung'), para(1, 'Ort und Datum'))
    const [range] = findPassages(document, 'zur Eignung')
    expect(textBetween(document, range)).toBe('zur Eignung')
  })

  it('finds a passage across a paragraph break and a line hyphenated in the PDF', () => {
    const document = doc(para(0, 'Der Bieter erklärt sein'), para(1, 'Angebot verbindlich.'))
    const [range] = findPassages(document, 'erklärt sein Ange-\nbot verbindlich')
    expect(textBetween(document, range)).toBe('erklärt sein\nAngebot verbindlich')
  })

  it('finds text in a table cell', () => {
    const document = doc(table(cell(para(0, 'Umsatz 2025')), cell(para(1, '1,2 Mio. €'))))
    expect(findPassages(document, 'Umsatz 2025 1,2 Mio. €')).toHaveLength(1)
  })

  it('finds every occurrence, in document order', () => {
    const document = doc(para(0, 'Ort, Datum'), para(1, 'Unterschrift'), para(2, 'Ort, Datum'))
    expect(findPassages(document, 'Ort, Datum').map((range) => range.from)).toEqual([1, 27])
  })

  it('finds nothing for text that is not there, or no text at all', () => {
    const document = doc(para(0, 'Ort, Datum'))
    expect(findPassages(document, 'Unterschrift')).toEqual([])
    expect(findPassages(document, ' - ')).toEqual([])
  })
})

describe('revealPassage', () => {
  it('marks the passage and scrolls it to the middle of the view', () => {
    const v = viewOf(para(0, 'Erklärung zur Eignung'))
    expect(revealPassage(v, { text: 'zur Eignung' }, () => 1)).toBe(true)
    expect(marked(v)).toBe('zur Eignung')
    expect(Element.prototype.scrollIntoView).toHaveBeenCalledWith(expect.objectContaining({ block: 'center' }))
  })

  it('takes the occurrence on the page nearest the hint', () => {
    const v = viewOf(para(0, 'Ort, Datum'), para(1, 'Unterschrift'), para(2, 'Ort, Datum'))
    const pageOf = (pos: number) => (pos < 20 ? 1 : 3)
    revealPassage(v, { text: 'Ort, Datum', pageHint: 3 }, pageOf)
    expect(revealKey.getState(v.state)!.find().map((mark) => mark.from)).toEqual([27])
    revealPassage(v, { text: 'Ort, Datum', pageHint: 1 }, pageOf)
    expect(revealKey.getState(v.state)!.find().map((mark) => mark.from)).toEqual([1])
  })

  it('marks one passage at a time and clears the mark when the text is not there', () => {
    const v = viewOf(para(0, 'Erklärung zur Eignung'), para(1, 'Ort und Datum'))
    revealPassage(v, { text: 'Eignung' }, () => 1)
    revealPassage(v, { text: 'Ort und Datum' }, () => 1)
    expect(marked(v)).toBe('Ort und Datum')
    expect(revealPassage(v, { text: 'Unterschrift' }, () => 1)).toBe(false)
    expect(marked(v)).toBe('')
  })
})

describe('pageOfNode', () => {
  it('counts the page gaps laid out before the node', () => {
    const root = document.createElement('div')
    root.innerHTML = '<p id="a">a</p><div class="page-gap"></div><p id="b">b</p><div class="page-gap"></div><p id="c">c</p>'
    expect(['a', 'b', 'c'].map((id) => pageOfNode(root.querySelector(`#${id}`)!, root))).toEqual([1, 2, 3])
  })
})

describe('createRevealQueue', () => {
  const request = (requestId: string, text: string) => ({ type: 'vound:reveal', requestId, text }) as const
  const result = (requestId: string, found: boolean) => ({ type: 'vound:revealResult', requestId, found })

  it('answers a reveal at once while no open is loading', () => {
    const post = vi.fn()
    createRevealQueue(post).reveal(viewOf(para(0, 'Ort und Datum')), request('r1', 'Datum'))
    expect(post.mock.calls).toEqual([[result('r1', true)]])
  })

  it('tells the host the opened document is in, once, then answers the reveal that waited for it', () => {
    const post = vi.fn()
    const queue = createRevealQueue(post)
    const v = viewOf(para(0, 'previous document'))
    queue.opening()
    queue.reveal(v, request('r1', 'Datum'))
    queue.reveal(v, request('r2', 'Ort'))
    expect(post.mock.calls).toEqual([[result('r1', false)]])

    v.dispatch(v.state.tr.replaceWith(0, v.state.doc.content.size, para(0, 'Ort und Datum')))
    queue.documentChanged(v)
    queue.documentChanged(v)
    expect(post.mock.calls).toEqual([[result('r1', false)], [{ type: 'vound:opened' }], [result('r2', true)]])
  })

  it('answers a reveal that waited for an open the next open replaced with not found', () => {
    const post = vi.fn()
    const queue = createRevealQueue(post)
    queue.opening()
    queue.reveal(null, request('r1', 'Datum'))
    queue.opening()
    expect(post).toHaveBeenCalledWith(result('r1', false))
  })
})
