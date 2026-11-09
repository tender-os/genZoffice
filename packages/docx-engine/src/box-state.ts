// Vound fork: writing a paragraph's tick boxes back into its original XML. Word draws both kinds from their state,
// so a tick has to reach the state: w:checked of a legacy FORMCHECKBOX, w14:checked of a checkbox content control,
// whose w:t holds the glyph as well.
import { isOn, sdtCheckboxIsChecked, syncFormCheckbox, syncSdtCheckbox } from './checkbox-control'
import type { Run } from './types'
import { escapeXmlText } from './xml-utils'

export type BoxRun = Pick<Run, 'text' | 'fldBeginXml' | 'sdtCheckboxXml'>
type Range = { start: number; end: number }

const LEGACY_BOX = /<w:checkBox\b(?:[^>]*\/>|[\s\S]*?<\/w:checkBox>)/g
const CHECKED_GLYPHS = ['☒', '☑']
const SDT_OPENS = ['<w:sdt>', '<w:sdt ']
const SDT_CLOSE = '</w:sdt>'
const CONTROL_MARK = '<w14:checkbox'

/** The offsets of each checkbox content control in the paragraph XML. */
export function checkboxControlRanges(paraXml: string): Range[] {
  const ranges: Range[] = []
  for (
    let at = paraXml.indexOf(CONTROL_MARK);
    at !== -1;
    at = paraXml.indexOf(CONTROL_MARK, at + 1)
  ) {
    const start = Math.max(...SDT_OPENS.map((open) => paraXml.lastIndexOf(open, at)))
    const close = paraXml.indexOf(SDT_CLOSE, at)
    if (start !== -1 && close !== -1) ranges.push({ start, end: close + SDT_CLOSE.length })
  }
  return ranges
}

/** The glyphs of the legacy boxes the runs show, one per box. */
const legacyGlyphs = (runs: ReadonlyArray<BoxRun>) =>
  runs.filter((run) => run.fldBeginXml).flatMap((run) => [...run.text])

/** The glyph of each checkbox control the runs show. */
const controlGlyphs = (runs: ReadonlyArray<BoxRun>) =>
  runs.filter((run) => run.sdtCheckboxXml).map((run) => [...run.text][0] ?? '')

export const legacyBoxCount = (paraXml: string) => paraXml.match(LEGACY_BOX)?.length ?? 0

/** Whether the runs still show every box of the paragraph. */
export const showsEveryBox = (paraXml: string, runs: ReadonlyArray<BoxRun>) =>
  legacyGlyphs(runs).length === legacyBoxCount(paraXml) &&
  controlGlyphs(runs).length === checkboxControlRanges(paraXml).length

function legacyBoxChecked(box: string) {
  const state = /<w:checked\b[^>]*>/.exec(box) ?? /<w:default\b[^>]*>/.exec(box)
  return state !== null && isOn(/w:val="([^"]*)"/.exec(state[0])?.[1])
}

function withLegacyStates(paraXml: string, glyphs: string[]) {
  let index = 0
  return paraXml.replace(LEGACY_BOX, (box) => {
    const checked = CHECKED_GLYPHS.includes(glyphs[index++])
    return checked === legacyBoxChecked(box) ? box : syncFormCheckbox(box, checked)
  })
}

function controlChecked(sdtPr: string) {
  const state = /<w14:checked\b[^>]*\bw14:val="([^"]*)"/.exec(sdtPr)
  return state !== null && isOn(state[1])
}

/** One checkbox control with its state and its w:t set from the glyph. */
function withControlState(control: string, glyph: string) {
  const sdtPrEnd = control.indexOf('</w:sdtPr>')
  const sdtPrStart = control.indexOf('<w:sdtPr')
  if (sdtPrStart === -1 || sdtPrEnd === -1) return control
  const sdtPr = control.slice(sdtPrStart, sdtPrEnd + '</w:sdtPr>'.length)
  const checked = sdtCheckboxIsChecked(sdtPr, glyph)
  const stated = checked === controlChecked(sdtPr) ? sdtPr : syncSdtCheckbox(sdtPr, checked)
  const content = control
    .slice(sdtPrEnd + '</w:sdtPr>'.length)
    .replace(
      /(<w:t(?:\s[^>]*)?>)[^<]*(<\/w:t>)/,
      (_, open: string, close: string) => open + escapeXmlText(glyph) + close,
    )
  return control.slice(0, sdtPrStart) + stated + content
}

function withControlStates(paraXml: string, glyphs: string[]) {
  const ranges = checkboxControlRanges(paraXml)
  return ranges.reduceRight(
    (xml, { start, end }, index) =>
      xml.slice(0, start) + withControlState(xml.slice(start, end), glyphs[index]) + xml.slice(end),
    paraXml,
  )
}

/** The paragraph XML with each box's state taken from the runs; only boxes whose state differs change. */
export const withBoxStates = (paraXml: string, runs: ReadonlyArray<BoxRun>) =>
  withControlStates(withLegacyStates(paraXml, legacyGlyphs(runs)), controlGlyphs(runs))
