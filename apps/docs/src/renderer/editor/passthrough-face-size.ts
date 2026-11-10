// Vound fork: the face and size the host's Format control gave a field passthrough paragraph (web-full
// fieldStyleBlock.ts), written into the paragraph's runs on save: the restyled runs' own, and the display's on the
// legacy text form field results, for text typed into the field after it.
import {
  withRestyledRuns,
  withResultFaceAndSize,
  type Block,
  type FieldDisplay,
} from '@genoffice/docx-engine'
import type { PmNode } from './convert'

const runFaces = (display: FieldDisplay) =>
  JSON.stringify(
    (display.runs ?? []).map((run) => [
      run.text,
      run.fontAscii ?? null,
      run.sizeHalfPoints ?? null,
    ]),
  )

const displayFace = (display: FieldDisplay) =>
  JSON.stringify([display.fontFamily ?? null, display.szHalfPoints ?? null])

/** the rewrite of the paragraph's XML for a restyled passthrough, or null when its face and size are as loaded */
export function passthroughFaceAndSize(
  node: PmNode,
  original: Block,
): ((xml: string) => string) | null {
  const current = node.attrs?.fieldDisplay as FieldDisplay | null
  const initial = original.fieldDisplay
  if (!current || !initial) return null
  const runsRestyled = current.runs !== undefined && runFaces(current) !== runFaces(initial)
  const displayRestyled = displayFace(current) !== displayFace(initial)
  if (!runsRestyled && !displayRestyled) return null
  return (xml) => {
    const withRuns = runsRestyled && current.runs ? withRestyledRuns(xml, current.runs) : xml
    if (!displayRestyled) return withRuns
    return withResultFaceAndSize(withRuns, {
      fontAscii: current.fontFamily,
      sizeHalfPoints: current.szHalfPoints,
    })
  }
}
