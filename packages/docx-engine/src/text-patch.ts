import { isOn } from './checkbox-control'
import { decodeEntities } from './parse-xml-text'
import type { Run } from './types'
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
      const patched = patchOneParagraph(para.xml, newParaTexts[i], skipLeading)
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
  /** offset of the opening <w:t ...> tag */
  tagStart: number
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
      innerStart,
      innerEnd: innerStart + inner.length,
      text: decodeEntities(inner),
    })
  }
  return out
}

function decodedTextOf(paraXml: string): string {
  return tSlices(paraXml)
    .map((t) => t.text)
    .join('')
}

function patchOneParagraph(paraXml: string, newText: string, skipLeading: boolean): string | null {
  const slices = tSlices(paraXml)
  if (slices.length === 0) return null
  const raw = slices.map((t) => t.text).join('')
  // First paragraph's leading whitespace (spacer run after the self-reference mark) is
  // excluded from the diff and treated as an immutable prefix
  const lead = skipLeading ? raw.length - raw.replace(/^\s+/, '').length : 0
  const oldText = raw.slice(lead)
  if (oldText === newText) return paraXml

  let prefix = 0
  const maxCommon = Math.min(oldText.length, newText.length)
  while (prefix < maxCommon && oldText[prefix] === newText[prefix]) prefix++
  let suffix = 0
  while (
    suffix < maxCommon - prefix &&
    oldText[oldText.length - 1 - suffix] === newText[newText.length - 1 - suffix]
  ) {
    suffix++
  }
  // Changed range (relative to the full raw text)
  const changeStart = lead + prefix
  const changeEnd = lead + oldText.length - suffix
  const replacement = newText.slice(prefix, newText.length - suffix)

  // Find the w:t intersecting the changed range; for pure insertion (changeStart ===
  // changeEnd), take the slice covering that point
  let acc = 0
  let first = -1
  let last = -1
  const bounds: Array<{ start: number; end: number }> = []
  for (let i = 0; i < slices.length; i++) {
    const start = acc
    const end = acc + slices[i].text.length
    bounds.push({ start, end })
    const touches =
      changeStart === changeEnd
        ? start <= changeStart && changeStart <= end
        : start < changeEnd && end > changeStart
    if (touches) {
      if (first === -1) first = i
      last = i
    }
    acc = end
  }
  if (first === -1) {
    // Insertion point falls outside every w:t (in theory only an empty paragraph) —
    // no anchor, give up
    return null
  }

  // Rewrite [first..last]: first takes the prefix remainder + replacement text + last's
  // suffix remainder; the rest are emptied
  const newInner: string[] = []
  for (let i = first; i <= last; i++) {
    if (i === first) {
      const head = slices[i].text.slice(0, changeStart - bounds[i].start)
      const tail = slices[last].text.slice(changeEnd - bounds[last].start)
      newInner.push(escapeXmlText(head + replacement + tail))
    } else {
      newInner.push('')
    }
  }

  // Replace back-to-front so offsets stay valid; changed w:t all get xml:space="preserve"
  let out = paraXml
  for (let i = last; i >= first; i--) {
    const s = slices[i]
    const closeEnd = s.innerEnd + '</w:t>'.length
    out =
      out.slice(0, s.tagStart) +
      `<w:t xml:space="preserve">${newInner[i - first]}</w:t>` +
      out.slice(closeEnd)
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
 * text cannot be mapped (tabs, breaks, checkbox glyphs, formatting-only edits
 * fall back to the regular rebuild).
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
  // not add or remove one, they stay where they are and the w:t mapping ignores them
  const controls = (joined.match(/[\t\n]/g) ?? []).length
  const originalControls = (originalParaXml.match(/<w:tab\s*\/>|<w:tab>\s*<\/w:tab>|<w:br\s*\/>|<w:cr\s*\/>/g) ?? []).length
  if (controls !== originalControls) return null
  const newText = joined.replace(/[\t\n]/g, '')
  const patched = patchParagraphTexts(originalParaXml, newText)
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
    return sdt
      .replace(/<w:showingPlcHdr\s*\/>/g, '')
      .replace(/<w:rStyle w:val="(?:[^"]*Platzhalter[^"]*|[^"]*Placeholder[^"]*)"\s*\/>/gi, '')
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
