import { decodeEntities } from './parse-xml-text'

const FORM_TEXT = 'FORMTEXT'
const FIELD_BEGIN = 'w:fldCharType="begin"'
const FIELD_SEPARATE = 'w:fldCharType="separate"'
const FIELD_END = 'w:fldCharType="end"'
const RUN_CLOSE = '</w:r>'
const HIGHLIGHT_RE = /<w:highlight\b[^>]*\/>/g
const TEXT_RE = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g
const DEFAULT_RE = /<w:textInput>[\s\S]*?<w:default w:val="([^"]*)"/

const runStartBefore = (xml: string, at: number) =>
  Math.max(xml.lastIndexOf('<w:r>', at), xml.lastIndexOf('<w:r ', at))

const textIn = (xml: string) =>
  [...xml.matchAll(TEXT_RE)].map((match) => decodeEntities(match[1])).join('')

type FieldSpan = { from: number; to: number; holdsValue: boolean }

/** each legacy text field of a paragraph, its begin run to its end run, and whether its result is a value */
function formTextFields(paraXml: string): FieldSpan[] {
  const fields: FieldSpan[] = []
  for (
    let code = paraXml.indexOf(FORM_TEXT);
    code !== -1;
    code = paraXml.indexOf(FORM_TEXT, code + 1)
  ) {
    const begin = paraXml.lastIndexOf(FIELD_BEGIN, code)
    const separate = paraXml.indexOf(FIELD_SEPARATE, code)
    if (begin === -1 || separate === -1) continue
    const end = paraXml.indexOf(FIELD_END, separate)
    const to = end === -1 ? paraXml.length : paraXml.indexOf(RUN_CLOSE, end) + RUN_CLOSE.length
    const result = textIn(paraXml.slice(separate, end === -1 ? paraXml.length : end))
    const shownDefault = DEFAULT_RE.exec(paraXml.slice(begin, separate))?.[1]
    const holdsValue = result.trim() !== '' && result !== decodeEntities(shownDefault ?? '')
    fields.push({ from: runStartBefore(paraXml, begin), to, holdsValue })
  }
  return fields
}

/** Vound fork: a legacy text field holding a value loses the author's fill-here w:highlight, begin to end */
export function withoutFilledFieldHighlight(paraXml: string): string {
  return formTextFields(paraXml)
    .filter((field) => field.holdsValue)
    .reverse()
    .reduce(
      (xml, field) =>
        xml.slice(0, field.from) +
        xml.slice(field.from, field.to).replace(HIGHLIGHT_RE, '') +
        xml.slice(field.to),
      paraXml,
    )
}
