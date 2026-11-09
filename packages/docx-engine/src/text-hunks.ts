import { lcsPairs } from './lcs-pairs'

/** a change to the old text: old [from, to) becomes `text` */
export interface TextHunk {
  from: number
  to: number
  text: string
}

type Stretch = { start: number; end: number }

/** whether the changes around old text [from, to) both lie in one stretch that holds more than that text */
const sameStretch = (previous: TextHunk, next: TextHunk, stretches: Stretch[]) =>
  stretches.some(
    ({ start, end }) =>
      start <= previous.from && next.to <= end && (start < previous.to || next.from < end),
  )

/** The changes with those of one stretch of the old text (a w:t) joined: a value retyped over its old
 *  characters is one replacement, not an insertion at the stretch's edge and a deletion inside it. */
export function joinedWithinStretches(
  hunks: TextHunk[],
  oldText: string,
  stretches: Stretch[],
): TextHunk[] {
  const joined: TextHunk[] = []
  for (const hunk of hunks) {
    const previous = joined.at(-1)
    if (previous && sameStretch(previous, hunk, stretches)) {
      const between = oldText.slice(previous.to, hunk.from)
      joined[joined.length - 1] = {
        from: previous.from,
        to: hunk.to,
        text: previous.text + between + hunk.text,
      }
    } else joined.push(hunk)
  }
  return joined
}

/** ponytail: the character diff is quadratic; a changed stretch longer than this stays one hunk */
const MAX_DIFF_CELLS = 250_000

/** The separate changes that turn `oldText` into `newText`, in order: two edited fields of one paragraph stay
 *  two changes instead of one stretch over the text between them. */
export function textHunks(oldText: string, newText: string): TextHunk[] {
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
  const oldMiddle = oldText.slice(prefix, oldText.length - suffix)
  const newMiddle = newText.slice(prefix, newText.length - suffix)
  if (oldMiddle.length * newMiddle.length > MAX_DIFF_CELLS) {
    return [{ from: prefix, to: prefix + oldMiddle.length, text: newMiddle }]
  }
  const pairs = lcsPairs(oldMiddle.split(''), newMiddle.split(''), (a, b) => a === b)
  const hunks: TextHunk[] = []
  let oldAt = 0
  let newAt = 0
  for (const [oldIndex, newIndex] of [...pairs, [oldMiddle.length, newMiddle.length]]) {
    if (oldIndex > oldAt || newIndex > newAt) {
      hunks.push({
        from: prefix + oldAt,
        to: prefix + oldIndex,
        text: newMiddle.slice(newAt, newIndex),
      })
    }
    oldAt = oldIndex + 1
    newAt = newIndex + 1
  }
  return hunks
}
