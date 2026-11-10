import { formTextResultSpans } from './form-text-anchor'
import type { Run } from './types'
import { escapeXmlText } from './xml-utils'

type ResultRun = { start: number; end: number }

const RUN_RE = /<w:r(?=[\s>])[\s\S]*?<\/w:r>/g

/** the runs with text inside a legacy text field's result */
function resultTextRuns(xml: string, span: { from: number; to: number }): ResultRun[] {
  RUN_RE.lastIndex = span.from
  const runs: ResultRun[] = []
  for (let match = RUN_RE.exec(xml); match && match.index < span.to; match = RUN_RE.exec(xml)) {
    if (/<w:t[\s>]/.test(match[0]))
      runs.push({ start: match.index, end: match.index + match[0].length })
  }
  return runs
}

const rawAttr = (rPr: string, tag: string, attr: string) =>
  new RegExp(`<${tag}\\b[^>]*\\s${attr}="([^"]*)"`).exec(rPr)?.[1]

/** the run shows another face or size than its run properties in the file say */
function restyled(run: Run) {
  const rPr = run.rawRPr ?? ''
  const face = rawAttr(rPr, 'w:rFonts', 'w:ascii') ?? run.themeRFonts?.fontAscii
  const size = rawAttr(rPr, 'w:sz', 'w:val')
  return (
    (run.fontAscii ?? undefined) !== face ||
    (run.sizeHalfPoints ?? undefined) !== (size === undefined ? undefined : Number(size))
  )
}

const runXml = (run: Run, rPrOf: (run: Run) => string) =>
  `<w:r>${rPrOf(run)}<w:t xml:space="preserve">${escapeXmlText(run.text.replace(/[\t\n]/g, ''))}</w:t></w:r>`

/** Vound fork: a legacy text field's result written as the field's runs in the editor, one run per formatting,
 *  once the face or size of one of them changed; any other result keeps its bytes. `rPrOf` is the
 *  generator's run properties of a run (mergeRPrModel). */
export function withFormTextStyles(
  xml: string,
  runs: ReadonlyArray<Run>,
  rPrOf: (run: Run) => string,
): string {
  const spans = formTextResultSpans(xml)
  let out = xml
  for (let index = spans.length - 1; index >= 0; index--) {
    const typed = runs.filter((run) => run.formText === index && run.text !== '')
    const inside = resultTextRuns(out, spans[index])
    if (inside.length === 0 || !typed.some(restyled)) continue
    const first = inside[0]
    const last = inside[inside.length - 1]
    out =
      out.slice(0, first.start) +
      typed.map((run) => runXml(run, rPrOf)).join('') +
      out.slice(last.end)
  }
  return out
}
