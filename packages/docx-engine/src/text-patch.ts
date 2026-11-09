import { isOn } from './checkbox-control'
import { decodeEntities } from './parse-xml-text'
import { withoutPlaceholderFlag, withoutPlaceholderStyle } from './parse-sdt'
import type { Run } from './types'
import { textHunks } from './text-hunks'
import { escapeXmlText } from './xml-utils'

/**
 * Paragraph-level surgical text patch (saving edits to rich-text entries such as
 * comments/footnotes).
 *
 * Maps the edited plain text ('\n' between paragraphs, each paragraph = concatenated w:t)
 * back onto the original entry XML:
 * - Paragraphs whose text is unchanged keep their original bytes (inline formatting,
 *   hyperlinks, images, and fields are all preserved)
 * - Changed paragraphs get a minimal w:t replacement: only the w:t outside the common
 *   prefix/suffix are rewritten, other run bytes stay untouched
 * - If the paragraph count changed, or a changed paragraph has no w:t to anchor to,
 *   returns null and the caller falls back to rebuilding the whole entry
 *
 * stripFirstParaLeadingSpace: footnote/endnote first paragraphs start with a
 * self-reference mark + space run, which the plain-text model drops; when enabled,
 * that leading whitespace is treated as an immutable prefix.
 */
export function patchParagraphTexts(
  entryXml: string,
  newText: string,
  opts: { stripFirstParaLeadingSpace?: boolean } = {},
): string | null {
  const paras = paragraphSlices(entryXml)
  if (paras.length === 0) return null
  const oldParaTexts = paras.map((p, i) =>
    i === 0 && opts.stripFirstParaLeadingSpace
      ? decodedTextOf(p.xml).replace(/^\s+/, '')
      : decodedTextOf(p.xml),
  )
  if (oldParaTexts.join('\n') === newText) return entryXml
  const newParaTexts = newText.split('\n')
  if (newParaTexts.length !== paras.length) return null

  let out = ''
  let cursor = 0
  for (let i = 0; i < paras.length; i++) {
    const para = paras[i]
    out += entryXml.slice(cursor, para.start)
    if (oldParaTexts[i] === newParaTexts[i]) {
      out += para.xml
    } else {
      const skipLeading = i === 0 && opts.stripFirstParaLeadingSpace === true
      const patched = patchOneParagraph(para.xml, newParaTexts[i], skipLeading, 'outsideText')
      if (patched === null) return null
      out += patched
    }
    cursor = para.end
  }
  out += entryXml.slice(cursor)

  // Self-check: re-extracted text must match the target, otherwise abandon the surgical patch
  const check = paragraphSlices(out).map((p, i) =>
    i === 0 && opts.stripFirstParaLeadingSpace
      ? decodedTextOf(p.xml).replace(/^\s+/, '')
      : decodedTextOf(p.xml),
  )
  return check.join('\n') === newText ? out : null
}

interface ParaSlice {
  start: number
  end: number
  xml: string
}

function paragraphSlices(xml: string): ParaSlice[] {
  // depth-aware scan: a paragraph may legally nest more <w:p> inside embedded
  // textbox content (w:txbxContent), so a lazy open-to-first-close match would
  // truncate the outer paragraph
  const out: ParaSlice[] = []
  const re = /<w:p(?:\s[^>]*)?\/>|<w:p(?:\s[^>]*)?>|<\/w:p>/g
  let m: RegExpExecArray | null
  let depth = 0
  let start = -1
  while ((m = re.exec(xml)) !== null) {
    const token = m[0]
    if (token === '</w:p>') {
      if (depth === 0) continue
      depth--
      if (depth === 0 && start !== -1) {
        const end = m.index + token.length
        out.push({ start, end, xml: xml.slice(start, end) })
        start = -1
      }
    } else if (token.endsWith('/>')) {
      if (depth === 0) out.push({ start: m.index, end: m.index + token.length, xml: token })
    } else {
      if (depth === 0) start = m.index
      depth++
    }
  }
  return out
}

interface TSlice {
  /** offsets of the inner content within the paragraph XML */
  innerStart: number
  innerEnd: number
  /** offset of the opening <w:t ...> tag, and of the end of the element */
  tagStart: number
  end: number
  text: string
}

function tSlices(paraXml: string): TSlice[] {
  const out: TSlice[] = []
  const re = /<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(paraXml)) !== null) {
    const inner = m[2]
    const innerStart = m.index + m[0].length - inner.length - '</w:t>'.length
    out.push({
      tagStart: m.index,
      end: m.index + m[0].length,
      innerStart,
      innerEnd: innerStart + inner.length,
      text: decodeEntities(inner),
    })
  }
  return out
}

/** Vound fork: the w:t of a paragraph and its line breaks (w:br, w:cr) between them, a break as a newline */
function lineSlices(paraXml: string): TSlice[] {
  const breaks: TSlice[] = []
  const re = /<w:br\s*\/>|<w:cr\s*\/>/g
  let m: RegExpExecArray | null
  while ((m = re.exec(paraXml)) !== null) {
    const at = m.index
    breaks.push({ tagStart: at, end: at + m[0].length, innerStart: at, innerEnd: at, text: '\n' })
  }
  return [...tSlices(paraXml), ...breaks].sort((a, b) => a.tagStart - b.tagStart)
}

const wtXml = (text: string) => `<w:t xml:space="preserve">${escapeXmlText(text)}</w:t>`

/** Vound fork: where a paragraph patch finds line breaks: outside the patched text (a field's, a comment's),
 *  or in it, written back as w:br (a content control's) */
type LineBreaks = 'outsideText' | 'inText'

const LINE_BREAKS: Record<
  LineBreaks,
  {
    slices: (paraXml: string) => TSlice[]
    write: (text: string) => string
    /** the characters and elements the w:t mapping leaves alone */
    outsideChars: RegExp
    outsideElements: RegExp
    patch: (paraXml: string, newText: string) => string | null
  }
> = {
  outsideText: {
    slices: tSlices,
    write: wtXml,
    outsideChars: /[\t\n]/g,
    outsideElements: /<w:tab\s*\/>|<w:tab>\s*<\/w:tab>|<w:br\s*\/>|<w:cr\s*\/>/g,
    patch: patchParagraphTexts,
  },
  inText: {
    slices: lineSlices,
    write: (text) =>
      text
        .split('\n')
        .map((part) => (part === '' ? '' : wtXml(part)))
        .join('<w:br/>'),
    outsideChars: /\t/g,
    outsideElements: /<w:tab\s*\/>|<w:tab>\s*<\/w:tab>/g,
    patch: patchParagraphLines,
  },
}

/** Vound fork: a field's paragraph, or one with typed breaks (column, page), keeps line breaks outside the patch */
function lineBreaksOf(paraXml: string): LineBreaks {
  if (paraXml.includes('<w:fldChar') || paraXml.includes('<w:br ')) return 'outsideText'
  return 'inText'
}

/** Vound fork: one paragraph's text, line breaks included, patched onto its XML */
function patchParagraphLines(paraXml: string, newText: string): string | null {
  const patched = patchOneParagraph(paraXml, newText, false, 'inText')
  if (patched === null) return null
  const check = lineSlices(patched)
    .map((slice) => slice.text)
    .join('')
  return check === newText ? patched : null
}

function decodedTextOf(paraXml: string): string {
  return tSlices(paraXml)
    .map((t) => t.text)
    .join('')
}

function patchOneParagraph(
  paraXml: string,
  newText: string,
  skipLeading: boolean,
  lineBreaks: LineBreaks,
): string | null {
  const { slices: slicesOf, write } = LINE_BREAKS[lineBreaks]
  const slices = slicesOf(paraXml)
  if (slices.length === 0) return null
  const raw = slices.map((t) => t.text).join('')
  // First paragraph's leading whitespace (spacer run after the self-reference mark) is
  // excluded from the diff and treated as an immutable prefix
  const lead = skipLeading ? raw.length - raw.replace(/^\s+/, '').length : 0
  const oldText = raw.slice(lead)
  if (oldText === newText) return paraXml

  const texts = slices.map((t) => t.text)
  const bounds: Array<{ start: number; end: number }> = []
  let acc = 0
  for (const slice of slices) {
    bounds.push({ start: acc, end: acc + slice.text.length })
    acc += slice.text.length
  }
  const rewritten = new Set<number>()
  // back to front: a later change never moves the offsets of an earlier one
  for (const hunk of textHunks(oldText, newText).reverse()) {
    const changeStart = lead + hunk.from
    const changeEnd = lead + hunk.to
    // the w:t the change touches; a pure insertion takes the first one covering its point
    const touched = bounds.flatMap(({ start, end }, i) => {
      const touches =
        changeStart === changeEnd
          ? start <= changeStart && changeStart <= end
          : start < changeEnd && end > changeStart
      return touches ? [i] : []
    })
    // insertion point outside every w:t (in theory only an empty paragraph): no anchor, give up
    if (touched.length === 0) return null
    const first = touched[0]
    const last = touched[touched.length - 1]
    // first takes its head, the new text and last's tail; the rest are emptied
    const head = texts[first].slice(0, changeStart - bounds[first].start)
    const tail = texts[last].slice(changeEnd - bounds[last].start)
    for (let i = first; i <= last; i++) {
      texts[i] = ''
      rewritten.add(i)
    }
    texts[first] = head + hunk.text + tail
  }

  // Replace back-to-front so offsets stay valid; changed w:t all get xml:space="preserve"
  let out = paraXml
  for (let i = slices.length - 1; i >= 0; i--) {
    if (!rewritten.has(i)) continue
    out = out.slice(0, slices[i].tagStart) + write(texts[i]) + out.slice(slices[i].end)
  }
  return out
}

/**
 * Vound fork: a paragraph that carries a complex field (w:fldChar, e.g. a legacy
 * FORMTEXT with its w:ffData) or a content control (w:sdt) must not be rebuilt
 * from the run model, because the model only holds the visible result text and
 * the rebuild drops the field machinery. When the edit only changed text, patch
 * the w:t inside the original paragraph XML instead: fields, controls, run
 * properties and the paragraph properties stay byte for byte. A placeholder
 * control that received real text loses its placeholder flag and style, as in
 * Word. Returns null when the paragraph has no field or control, or when the
 * text cannot be mapped (added or removed tabs and breaks around a plain field,
 * checkbox glyphs, formatting-only edits fall back to the regular rebuild). A form
 * field or control keeps its paragraph through any such edit, since a rebuild would
 * lose it: an added tab or break becomes a space, and a removed one stays.
 */
export function patchFieldParagraphText(
  originalParaXml: string,
  runs: ReadonlyArray<WtMappedRun>,
): string | null {
  if (!/<w:fldChar\b|<w:sdt\b/.test(originalParaXml)) return null
  if (/<w:noBreakHyphen\b|<w:softHyphen\b|<w:ruby\b/.test(originalParaXml)) return null
  if (!sameContentOutsideWt(originalParaXml, runs)) return null
  const joined = runs.map(wtTextOf).join('')
  if (/[\f\v\u001e]/.test(joined)) return null
  // tabs and line breaks are their own elements, not w:t text: as long as the edit did
  // not add or remove one, they stay where they are and the w:t mapping ignores them;
  // in a content control's paragraph a line break is written back as a w:br (Vound fork)
  const mode = LINE_BREAKS[lineBreaksOf(originalParaXml)]
  const controls = (joined.match(mode.outsideChars) ?? []).length
  const originalControls = (originalParaXml.match(mode.outsideElements) ?? []).length
  if (controls !== originalControls && !/<w:ffData\b|<w:sdt\b/.test(originalParaXml)) return null
  const newText = joined.replace(mode.outsideChars, controls > originalControls ? ' ' : '')
  const patched = mode.patch(originalParaXml, newText)
  if (patched === null) return null
  if (patched === originalParaXml || !patched.includes('<w:showingPlcHdr')) return patched
  // typed text replaced a placeholder: that control loses its placeholder flag and the
  // placeholder character style, as in Word; controls whose text is unchanged keep both
  const sdtRe = /<w:sdt>[\s\S]*?<\/w:sdt>/g
  const before = originalParaXml.match(sdtRe) ?? []
  let i = 0
  return patched.replace(sdtRe, (sdt) => {
    const prior = before[i++]
    if (prior !== undefined && decodedTextOf(prior) === decodedTextOf(sdt)) return sdt
    return withoutPlaceholderStyle(withoutPlaceholderFlag(sdt))
  })
}

type WtMappedRun = Pick<Run, 'text' | 'noteRef' | 'sym' | 'math' | 'del' | 'fldBeginXml'>

/** A run's text that lives in w:t: note markers, symbols, formulas, tracked deletions and legacy box glyphs do not. */
function wtTextOf(run: WtMappedRun): string {
  return run.noteRef || run.sym || run.math || run.del || run.fldBeginXml ? '' : run.text
}

/** Added, removed or changed content outside w:t needs the rebuild: a w:t patch would not reach it. */
function sameContentOutsideWt(paraXml: string, runs: ReadonlyArray<WtMappedRun>): boolean {
  const runCount = (kind: 'noteRef' | 'sym' | 'math') => runs.filter((r) => r[kind]).length
  const xmlCount = (element: RegExp) => paraXml.match(element)?.length ?? 0
  const runText = (kind: 'del' | 'fldBeginXml') =>
    runs
      .filter((r) => r[kind])
      .map((r) => r.text)
      .join('')
  const deletedText = [...paraXml.matchAll(/<w:delText(?:\s[^>]*)?>([\s\S]*?)<\/w:delText>/g)]
    .map((m) => decodeEntities(m[1]))
    .join('')
  return (
    runCount('noteRef') === xmlCount(/<w:(?:footnote|endnote)Reference\b/g) &&
    runCount('sym') === xmlCount(/<w:sym\b/g) &&
    runCount('math') === xmlCount(/<m:oMath\b/g) &&
    runText('del') === deletedText &&
    runText('fldBeginXml') === legacyBoxGlyphs(paraXml)
  )
}

/** The glyphs the parser shows for the paragraph's legacy checkboxes: w:checked wins over w:default. */
function legacyBoxGlyphs(paraXml: string): string {
  const boxes = paraXml.match(/<w:checkBox\b(?:[^>]*\/>|[\s\S]*?<\/w:checkBox>)/g) ?? []
  return boxes
    .map((box) => {
      const state = /<w:checked\b[^>]*>/.exec(box) ?? /<w:default\b[^>]*>/.exec(box)
      if (!state) return '☐'
      return isOn(/w:val="([^"]*)"/.exec(state[0])?.[1]) ? '☒' : '☐'
    })
    .join('')
}
