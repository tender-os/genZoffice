// Vound fork: a cell paragraph whose runs changed only in face or size, the host's Format control, is written by
// the cell text patch like an edited one, which keeps the cell's form fields.
import type { Run } from '@genoffice/docx-engine'

type FacedRun = Pick<Run, 'text' | 'fontAscii' | 'sizeHalfPoints'>

/** the face and size of each character of a paragraph */
const facesOf = (runs: ReadonlyArray<FacedRun>) =>
  runs
    .flatMap((run) => [...run.text].map(() => `${run.fontAscii ?? ''}|${run.sizeHalfPoints ?? ''}`))
    .join(',')

/** a paragraph's text shows another face or size than the original paragraph's */
export const facesChanged = (
  runs: ReadonlyArray<FacedRun>,
  original: ReadonlyArray<FacedRun> | undefined,
) => original !== undefined && facesOf(runs) !== facesOf(original)
