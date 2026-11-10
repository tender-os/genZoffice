// Vound fork: a field passthrough's formatted runs after the reviewer edited its text in place. The text before and
// after the edit keeps the run it had, and typed text takes the formatting of the character before it (the one
// after, at the start), as typing does in Word; so a face the host's Format control set on part of the line stays.
import type { FieldRun } from '@genoffice/docx-engine'

const commonSuffix = (a: string, b: string) => {
  let length = 0
  while (
    length < a.length &&
    length < b.length &&
    a[a.length - 1 - length] === b[b.length - 1 - length]
  )
    length++
  return length
}

const commonPrefix = (a: string, b: string, suffix: number) => {
  let length = 0
  while (length < a.length - suffix && length < b.length - suffix && a[length] === b[length])
    length++
  return length
}

/** the runs of `after`, or null when the runs do not describe `before` or hold a formula */
export function runsAfterEdit(
  runs: readonly FieldRun[],
  before: string,
  after: string,
): FieldRun[] | null {
  if (runs.map((run) => run.text).join('') !== before || runs.some((run) => run.math)) return null
  const owners = runs.flatMap((run, index) => [...run.text].map(() => index))
  if (owners.length === 0) return null
  // the suffix first: an edit next to a repeated character lands on its left, after the text it continues
  const suffix = commonSuffix(before, after)
  const prefix = commonPrefix(before, after, suffix)
  const typedOwner = owners[prefix > 0 ? prefix - 1 : Math.min(prefix, owners.length - 1)]
  const ownerAt = (offset: number) => {
    if (offset < prefix) return owners[offset]
    if (offset >= after.length - suffix) return owners[before.length - (after.length - offset)]
    return typedOwner
  }
  const next: FieldRun[] = []
  let last = -1
  ;[...after].forEach((char, offset) => {
    const owner = ownerAt(offset)
    if (owner === last)
      next[next.length - 1] = { ...next[next.length - 1], text: next[next.length - 1].text + char }
    else next.push({ ...runs[owner], text: char })
    last = owner
  })
  return next
}
