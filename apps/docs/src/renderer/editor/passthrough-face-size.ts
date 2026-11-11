// Vound fork: the face, size and alignment the host's Format control gave a field passthrough paragraph (web-full
// fieldStyleBlock.ts), written into the paragraph on save: the restyled runs' own face and size, the display's on
// the legacy text form field results, for text typed into the field after it, and the display's alignment as w:jc.
// A size the display or a run no longer has is taken off the runs that had it.
import {
  withParagraphAlignment,
  withRestyledRuns,
  withResultFaceAndSize,
  type Block,
  type FieldDisplay,
  type FieldRun,
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

/** the size each character of the runs has */
const sizesOf = (runs: readonly FieldRun[]) =>
  runs.flatMap((run) => [...run.text].map(() => run.sizeHalfPoints))

/** the current runs, a run that lost the size its text had before asking for none */
function withTakenSizes(current: readonly FieldRun[], initial: readonly FieldRun[]) {
  const before = sizesOf(initial)
  let offset = 0
  return current.map((run) => {
    const had = before[offset]
    offset += run.text.length
    return run.sizeHalfPoints === undefined && had !== undefined
      ? { ...run, sizeHalfPoints: null }
      : run
  })
}

/** the rewrite of the paragraph's XML for a restyled passthrough, or null when it is as loaded */
export function passthroughFaceAndSize(
  node: PmNode,
  original: Block,
): ((xml: string) => string) | null {
  const current = node.attrs?.fieldDisplay as FieldDisplay | null
  const initial = original.fieldDisplay
  if (!current || !initial) return null
  const runsRestyled = current.runs !== undefined && runFaces(current) !== runFaces(initial)
  const displayRestyled = displayFace(current) !== displayFace(initial)
  const realigned = current.align !== undefined && current.align !== initial.align
  if (!runsRestyled && !displayRestyled && !realigned) return null
  return (xml) => {
    const runs =
      runsRestyled && current.runs ? withTakenSizes(current.runs, initial.runs ?? []) : null
    const withRuns = runs ? withRestyledRuns(xml, runs) : xml
    const withDisplay = displayRestyled
      ? withResultFaceAndSize(withRuns, {
          fontAscii: current.fontFamily,
          sizeHalfPoints:
            current.szHalfPoints ?? (initial.szHalfPoints === undefined ? undefined : null),
        })
      : withRuns
    return realigned && current.align
      ? withParagraphAlignment(withDisplay, current.align)
      : withDisplay
  }
}
