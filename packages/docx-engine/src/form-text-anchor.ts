const FORM_TEXT = 'FORMTEXT'
const FIELD_END = 'w:fldCharType="end"'
const FIELD_SEPARATE = 'w:fldCharType="separate"'
const EMPTY_RESULT_RUN = '<w:r><w:t xml:space="preserve"></w:t></w:r>'

/** Vound fork: an empty w:t in the result of each legacy text field that has none, so text typed for the field
 *  has a place inside it */
export function withFormTextAnchors(paraXml: string): string {
  let out = ''
  let cursor = 0
  for (
    let field = paraXml.indexOf(FORM_TEXT);
    field !== -1;
    field = paraXml.indexOf(FORM_TEXT, cursor)
  ) {
    const end = paraXml.indexOf(FIELD_END, field)
    if (end === -1) break
    const endRun = Math.max(paraXml.lastIndexOf('<w:r>', end), paraXml.lastIndexOf('<w:r ', end))
    const result = paraXml.slice(field, endRun)
    const holdsText = result.includes('<w:t>') || result.includes('<w:t ')
    const separate = result.includes(FIELD_SEPARATE)
      ? ''
      : `<w:r><w:fldChar ${FIELD_SEPARATE}/></w:r>`
    out += paraXml.slice(cursor, endRun)
    if (!holdsText) out += separate + EMPTY_RESULT_RUN
    out += paraXml.slice(endRun, end)
    cursor = end
  }
  return out + paraXml.slice(cursor)
}

/** Vound fork: where each legacy text field's result lies, for XML that went through withFormTextAnchors */
export function formTextResultSpans(paraXml: string): { from: number; to: number }[] {
  const spans: { from: number; to: number }[] = []
  for (
    let field = paraXml.indexOf(FORM_TEXT);
    field !== -1;
    field = paraXml.indexOf(FORM_TEXT, field + 1)
  ) {
    const from = paraXml.indexOf(FIELD_SEPARATE, field)
    const to = from === -1 ? -1 : paraXml.indexOf(FIELD_END, from)
    if (to === -1) return []
    spans.push({ from, to })
  }
  return spans
}

/** Vound fork: where the result of a paragraph's only legacy text field lies; null without exactly one */
export function onlyFormTextResult(paraXml: string): { from: number; to: number } | null {
  const field = paraXml.indexOf(FORM_TEXT)
  if (field === -1 || paraXml.indexOf(FORM_TEXT, field + 1) !== -1) return null
  const from = paraXml.indexOf(FIELD_SEPARATE, field)
  const to = from === -1 ? -1 : paraXml.indexOf(FIELD_END, from)
  return to === -1 ? null : { from, to }
}
