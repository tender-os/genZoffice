/** Vound fork: the w:t texts after an edit of their joined text: the text before and after the change stays in its
 *  w:t, and the changed part goes into the w:t it starts in (an insertion between two goes into the left one, after
 *  the text it continues). Null when the joined text did not change. */
export function editedNodeTexts(value: string, texts: readonly string[]): string[] | null {
  const before = texts.join('')
  if (before === value) return null
  let suffix = 0
  while (
    suffix < before.length &&
    suffix < value.length &&
    before[before.length - 1 - suffix] === value[value.length - 1 - suffix]
  )
    suffix++
  let prefix = 0
  while (
    prefix < before.length - suffix &&
    prefix < value.length - suffix &&
    before[prefix] === value[prefix]
  )
    prefix++
  const ends = texts.map((_, i) => texts.slice(0, i + 1).join('').length)
  const starts = ends.map((end, i) => end - texts[i].length)
  const changedEnd = before.length - suffix
  const first = Math.max(
    0,
    starts.findIndex((start, i) => prefix >= start && prefix <= ends[i]),
  )
  const lastFound = ends.findIndex(
    (end, i) => changedEnd >= starts[i] && changedEnd <= end && i >= first,
  )
  const last = lastFound === -1 ? first : lastFound
  const inserted = value.slice(prefix, value.length - suffix)
  return texts.map((text, i) => {
    if (i < first || i > last) return text
    const head = i === first ? text.slice(0, prefix - starts[i]) : ''
    const tail = i === last ? text.slice(changedEnd - starts[i]) : ''
    return i === first ? head + inserted + (first === last ? tail : '') : tail
  })
}
