// Reveal (vound:reveal): the host names a passage it cites, the editor finds it in the document's text,
// scrolls it to the middle of the view and marks it, one passage at a time (the next reveal or open clears
// it). Matching ignores case, whitespace, hyphens and the quote style, and folds ligatures (NFKC), so a
// passage quoted from the converted PDF (line breaks, hyphenated words, typographic quotes, ﬁ) still finds
// the Word text, also across paragraph breaks. Several matches are told apart by the host's page hint.
// Display only: the mark is a decoration and never reaches the saved bytes.
import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import { glideIntoView } from './markersPlugin'

const QUOTES: Record<string, string> = {
  '„': '"', '“': '"', '”': '"', '‟': '"', '«': '"', '»': '"', '‚': "'", '‘': "'", '’': "'", '‛': "'",
}
/** hyphens and dashes a line break or a typesetter adds or swaps, and the soft hyphen */
const HYPHENS = new Set(['-', '\u00ad', '\u2010', '\u2011', '\u2012', '\u2013', '\u2014'])
const isIgnored = (c: string) => HYPHENS.has(c) || /\s|\u200b/.test(c)

/** one source character as the characters it is matched by (none for whitespace and hyphens) */
export function matchChars(ch: string): string[] {
  return [...ch.normalize('NFKC').toLowerCase()].map((c) => QUOTES[c] ?? c).filter((c) => !isIgnored(c))
}

export const matchText = (text: string) => [...text].flatMap(matchChars).join('')

/** the document's text as matched, with each matched character's document range */
function indexDocument(doc: PmNode) {
  let text = ''
  const from: number[] = []
  const to: number[] = []
  doc.descendants((node, pos) => {
    if (!node.isText) return true
    let offset = 0
    for (const ch of node.text ?? '') {
      for (const c of matchChars(ch)) {
        text += c
        from.push(pos + offset)
        to.push(pos + offset + ch.length)
      }
      offset += ch.length
    }
    return false
  })
  return { text, from, to }
}

export type PassageRange = { from: number; to: number }

/** every place the passage's text stands in the document, in document order */
export function findPassages(doc: PmNode, passage: string): PassageRange[] {
  const needle = matchText(passage)
  if (!needle) return []
  const index = indexDocument(doc)
  const found: PassageRange[] = []
  for (let at = index.text.indexOf(needle); at !== -1; at = index.text.indexOf(needle, at + 1)) {
    found.push({ from: index.from[at], to: index.to[at + needle.length - 1] })
  }
  return found
}

/** the match on the page nearest the hint; the first without one */
export function nearestToPage(ranges: PassageRange[], pageHint: number | undefined, pageOf: (pos: number) => number) {
  if (pageHint === undefined) return ranges[0]
  let best = ranges[0]
  for (const range of ranges) {
    if (Math.abs(pageOf(range.from) - pageHint) < Math.abs(pageOf(best.from) - pageHint)) best = range
  }
  return best
}

// ponytail: counts the page gaps the canvas draws before the node; check against rendered pages if hints miss
/** the 1-based page a DOM node is laid out on: one more than the page gaps before it */
export function pageOfNode(node: Node, root: Element): number {
  let page = 1
  for (const gap of root.querySelectorAll('.page-gap')) {
    if (gap.compareDocumentPosition(node) & Node.DOCUMENT_POSITION_FOLLOWING) page += 1
  }
  return page
}

const pageOfPos = (view: EditorView) => (pos: number) => pageOfNode(view.domAtPos(pos).node, view.dom)

export const revealKey = new PluginKey<DecorationSet>('voundReveal')

/** the revealed passage's mark; `onDocumentChanged` hears every new document, the opened one included */
export const revealPlugin = (onDocumentChanged: (view: EditorView) => void) =>
  new Plugin<DecorationSet>({
    key: revealKey,
    state: {
      init: () => DecorationSet.empty,
      apply: (tr, set) => tr.getMeta(revealKey) ?? set.map(tr.mapping, tr.doc),
    },
    props: { decorations: (state) => revealKey.getState(state) },
    view: () => ({ update: (view, previous) => { if (view.state.doc !== previous.doc) onDocumentChanged(view) } }),
  })

/** the class host.css gives the revealed passage */
export const REVEAL_CLASS = 'vound-reveal'

/** Marks and scrolls to the passage, replacing the previous mark; false, with no mark, when the text is not there. */
export function revealPassage(
  view: EditorView,
  request: { text: string; pageHint?: number },
  pageOf: (pos: number) => number = pageOfPos(view),
): boolean {
  const ranges = findPassages(view.state.doc, request.text)
  const range = ranges.length > 0 ? nearestToPage(ranges, request.pageHint, pageOf) : undefined
  const marks = range ? [Decoration.inline(range.from, range.to, { class: REVEAL_CLASS })] : []
  view.dispatch(view.state.tr.setMeta(revealKey, DecorationSet.create(view.state.doc, marks)))
  if (!range) return false
  const { node } = view.domAtPos(range.from)
  const element = node instanceof HTMLElement ? node : node.parentElement
  if (element) glideIntoView(element, 'center')
  return true
}

/** a new document opens: the previous passage's mark goes */
export function clearReveal(view: EditorView) {
  view.dispatch(view.state.tr.setMeta(revealKey, DecorationSet.empty))
}
