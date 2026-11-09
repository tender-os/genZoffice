import { decodeEntities } from './parse-xml-text'
import { withoutPlaceholderFlag, withoutPlaceholderStyle } from './parse-sdt'
import { EMPTY_CONTROL_TEXT, type Run } from './types'
import { checkboxControlRanges, showsEveryBox, withBoxStates } from './box-state'
import { formTextResultSpans, withFormTextAnchors } from './form-text-anchor'
import { joinedWithinStretches, textHunks } from './text-hunks'
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
    /** the paragraph XML the patch starts from */
    anchor: (paraXml: string) => string
  }
> = {
  outsideText: {
    slices: tSlices,
    write: wtXml,
    outsideChars: /[\t\n]/g,
    outsideElements: /<w:tab\s*\/>|<w:tab>\s*<\/w:tab>|<w:br\s*\/>|<w:cr\s*\/>/g,
    patch: patchParagraphTexts,
    anchor: (paraXml) => withControlAnchors(withFormTextAnchors(paraXml)),
  },
  inText: {
    slices: lineSlices,
    // an emptied stretch keeps an empty w:t: the next save anchors typed text in it
    write: (text) =>
      text
        .split('\n')
        .map((part) => (part === '' ? '' : wtXml(part)))
        .join('<w:br/>') || wtXml(''),
    outsideChars: /\t/g,
    outsideElements: /<w:tab\s*\/>|<w:tab>\s*<\/w:tab>/g,
    patch: patchParagraphLines,
    anchor: withControlAnchors,
  },
}

const CONTENT_OPEN = '<w:sdtContent>'
const CONTENT_CLOSE = '</w:sdtContent>'

/** Vound fork: an empty w:t in each control that holds none (emptied by an earlier save, or empty from the
 *  start), so text typed into it has a place inside the control */
function withControlAnchors(paraXml: string): string {
  let out = ''
  let cursor = 0
  for (let open = paraXml.indexOf(CONTENT_OPEN); open !== -1; open = paraXml.indexOf(CONTENT_OPEN, cursor)) {
    const start = open + CONTENT_OPEN.length
    const close = paraXml.indexOf(CONTENT_CLOSE, start)
    if (close === -1) break
    const content = paraXml.slice(start, close)
    const holdsText = content.includes('<w:t>') || content.includes('<w:t ') || content.includes(CONTENT_OPEN)
    const runEnd = content.indexOf('</w:r>')
    out += paraXml.slice(cursor, start)
    if (holdsText) out += content
    else if (runEnd === -1) out += `<w:r>${wtXml('')}</w:r>${content}`
    else out += content.slice(0, runEnd) + wtXml('') + content.slice(runEnd)
    cursor = close
  }
  return out + paraXml.slice(cursor)
}

/** Vound fork: a field's paragraph, or one with typed breaks (column, page), keeps line breaks outside the patch */
function lineBreaksOf(paraXml: string): LineBreaks {
  if (paraXml.includes('<w:fldChar') || paraXml.includes('<w:br ')) return 'outsideText'
  return 'inText'
}

const withoutFiller = (text: string) => text.replaceAll(EMPTY_CONTROL_TEXT, '')

/** Vound fork: one paragraph's text, line breaks included, patched onto its XML */
function patchParagraphLines(paraXml: string, newText: string): string | null {
  const patched = patchOneParagraph(paraXml, newText, false, 'inText')
  if (patched === null) return null
  const check = outsideCheckboxControls(patched, lineSlices(patched))
    .map((slice) => slice.text)
    .join('')
  return withoutFiller(check) === withoutFiller(newText) ? patched : null
}

function decodedTextOf(paraXml: string): string {
  return tSlices(paraXml)
    .map((t) => t.text)
    .join('')
}

/** Vound fork: whether an offset of the paragraph XML lies inside a content control's content */
const insideControl = (paraXml: string, at: number) =>
  paraXml.lastIndexOf(CONTENT_OPEN, at) > paraXml.lastIndexOf(CONTENT_CLOSE, at)

/** Vound fork: the slices text may go into: never a checkbox control's glyph */
function outsideCheckboxControls(paraXml: string, slices: TSlice[]) {
  const controls = checkboxControlRanges(paraXml)
  return slices.filter(({ tagStart }) => !controls.some(({ start, end }) => start <= tagStart && tagStart < end))
}

/** Vound fork: an empty run closing a paragraph no text can go into (only boxes), so typed text has a place */
function withClosingAnchor(paraXml: string, lineBreaks: LineBreaks) {
  if (outsideCheckboxControls(paraXml, LINE_BREAKS[lineBreaks].slices(paraXml)).length > 0) return paraXml
  const end = paraXml.lastIndexOf('</w:p>')
  return end === -1 ? paraXml : `${paraXml.slice(0, end)}<w:r>${wtXml('')}</w:r>${paraXml.slice(end)}`
}

function patchOneParagraph(
  paraXml: string,
  newText: string,
  skipLeading: boolean,
  lineBreaks: LineBreaks,
): string | null {
  const { slices: slicesOf, write } = LINE_BREAKS[lineBreaks]
  const slices = outsideCheckboxControls(paraXml, slicesOf(paraXml))
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
  const stretches = bounds.map(({ start, end }) => ({ start: start - lead, end: end - lead }))
  for (const hunk of joinedWithinStretches(textHunks(oldText, newText), oldText, stretches).reverse()) {
    const changeStart = lead + hunk.from
    const changeEnd = lead + hunk.to
    // the w:t the change touches; a pure insertion goes into one of those covering its point: an empty one
    // (a control emptied before), else one in a control, else the first
    const touched = bounds.flatMap(({ start, end }, i) => {
      const touches =
        changeStart === changeEnd
          ? start <= changeStart && changeStart <= end
          : start < changeEnd && end > changeStart
      return touches ? [i] : []
    })
    // insertion point outside every w:t (in theory only an empty paragraph): no anchor, give up
    if (touched.length === 0) return null
    const insertedInto =
      changeStart === changeEnd
        ? (touched.find((i) => bounds[i].start === bounds[i].end) ??
          touched.find((i) => insideControl(paraXml, slices[i].tagStart)) ??
          touched[0])
        : undefined
    const first = insertedInto ?? touched[0]
    const last = insertedInto ?? touched[touched.length - 1]
    // first takes its head, the new text and last's tail; the rest are emptied
    const head = texts[first].slice(0, changeStart - bounds[first].start)
    const tail = texts[last].slice(changeEnd - bounds[last].start)
    for (let i = first; i <= last; i++) {
      texts[i] = ''
      rewritten.add(i)
    }
    // a filler that lost its mark (it can only be that in a w:t that was empty) never reaches the file
    const intoEmpty = lineBreaks === 'inText' && bounds[first].start === bounds[first].end
    texts[first] = head + (intoEmpty ? withoutFiller(hunk.text) : hunk.text) + tail
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
  const lineBreaks = lineBreaksOf(originalParaXml)
  const mode = LINE_BREAKS[lineBreaks]
  const controls = (joined.match(mode.outsideChars) ?? []).length
  const originalControls = (originalParaXml.match(mode.outsideElements) ?? []).length
  if (controls !== originalControls && !/<w:ffData\b|<w:sdt\b/.test(originalParaXml)) return null
  const newText = joined.replace(mode.outsideChars, controls > originalControls ? ' ' : '')
  const boxed = withBoxStates(originalParaXml, runs)
  const source = withClosingAnchor(mode.anchor(boxed), lineBreaks)
  const byField = withFormTextValues(source, runs, mode.outsideChars)
  if (byField !== null) return byField
  const patched = mode.patch(source, newText)
  if (patched === null) return null
  if (patched === source) return boxed
  if (!patched.includes('<w:showingPlcHdr')) return patched
  // typed text replaced a placeholder: that control loses its placeholder flag and the
  // placeholder character style, as in Word; controls whose text is unchanged keep both
  const sdtRe = /<w:sdt>[\s\S]*?<\/w:sdt>/g
  const before = source.match(sdtRe) ?? []
  let i = 0
  return patched.replace(sdtRe, (sdt) => {
    const prior = before[i++]
    if (prior !== undefined && decodedTextOf(prior) === decodedTextOf(sdt)) return sdt
    return withoutPlaceholderStyle(withoutPlaceholderFlag(sdt))
  })
}

type WtMappedRun = Pick<
  Run,
  'text' | 'noteRef' | 'sym' | 'math' | 'del' | 'fldBeginXml' | 'placeholder' | 'sdtCheckboxXml' | 'formText'
>

/** Vound fork: each legacy text field's text in its own result; null when the text around the fields changed,
 *  which the mapping over the whole paragraph takes */
function withFormTextValues(
  xml: string,
  runs: ReadonlyArray<WtMappedRun>,
  outsideChars: RegExp,
): string | null {
  const results = formTextResultSpans(xml)
  if (!runs.some((run) => run.formText !== undefined)) return null
  if (runs.some((run) => (run.formText ?? 0) >= results.length)) return null
  const textOf = (matching: ReadonlyArray<WtMappedRun>) => withoutFiller(matching.map(wtTextOf).join(''))
  const slices = tSlices(xml)
  const insideResult = (slice: TSlice, result: { from: number; to: number }) =>
    slice.tagStart > result.from && slice.tagStart < result.to
  const around = slices.filter((slice) => !results.some((result) => insideResult(slice, result)))
  const typedAround = textOf(runs.filter((run) => run.formText === undefined)).replace(outsideChars, '')
  if (around.map((slice) => slice.text).join('') !== typedAround) return null
  let out = xml
  results.forEach((_, index) => {
    const last = results.length - 1 - index
    const typed = runs.filter((run) => run.formText === last)
    const inside = slices.filter((slice) => insideResult(slice, results[last]))
    const value = textOf(typed)
    if (typed.length === 0 || inside.map((slice) => slice.text).join('') === value) return
    for (let at = inside.length - 1; at >= 0; at--) {
      out = out.slice(0, inside[at].tagStart) + wtXml(at === 0 ? value : '') + out.slice(inside[at].end)
    }
  })
  return out
}

/** A run's text that lives in w:t: note markers, symbols, formulas, tracked deletions and legacy box glyphs do not. */
function wtTextOf(run: WtMappedRun): string {
  if (run.placeholder && run.text === EMPTY_CONTROL_TEXT) return ''
  // a checkbox control's first character is its glyph; anything after it was typed beside the control
  if (run.sdtCheckboxXml) return [...run.text].slice(1).join('')
  return run.noteRef || run.sym || run.math || run.del || run.fldBeginXml ? '' : run.text
}

/** Added, removed or changed content outside w:t needs the rebuild: a w:t patch would not reach it. */
function sameContentOutsideWt(paraXml: string, runs: ReadonlyArray<WtMappedRun>): boolean {
  const runCount = (kind: 'noteRef' | 'sym' | 'math') => runs.filter((r) => r[kind]).length
  const xmlCount = (element: RegExp) => paraXml.match(element)?.length ?? 0
  const runText = (kind: 'del') =>
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
    showsEveryBox(paraXml, runs)
  )
}
