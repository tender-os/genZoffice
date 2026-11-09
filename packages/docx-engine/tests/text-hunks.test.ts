import { describe, expect, test } from 'vitest'
import { textHunks, type TextHunk } from '../src/text-hunks'

const applied = (oldText: string, hunks: TextHunk[]) =>
  hunks.reduceRight(
    (text, hunk) => text.slice(0, hunk.from) + hunk.text + text.slice(hunk.to),
    oldText,
  )

/** a seeded generator, so a failure repeats */
function random(seed: number) {
  let state = seed
  return () => {
    state = (state * 1664525 + 1013904223) % 4294967296
    return state / 4294967296
  }
}

describe('textHunks', () => {
  test('two edited stretches stay two changes and leave the text between them alone', () => {
    // "Ort: " 0-5, "Ort eintragen" 5-18, ", Datum: " 18-27, "Datum eingeben" 27-41
    const hunks = textHunks(
      'Ort: Ort eintragen, Datum: Datum eingeben',
      'Ort: Berlin, Datum: 09.11.2026',
    )
    expect(hunks.every((hunk) => (hunk.from >= 5 && hunk.to <= 18) || hunk.from >= 27)).toBe(true)
  })

  test('applying the changes gives the new text, in order and without overlap', () => {
    const next = random(7)
    const alphabet = 'ab, :'
    const text = (length: number) =>
      Array.from({ length }, () => alphabet[Math.floor(next() * alphabet.length)]).join('')
    for (let run = 0; run < 500; run++) {
      const oldText = text(Math.floor(next() * 20))
      const newText = text(Math.floor(next() * 20))
      const hunks = textHunks(oldText, newText)
      expect(applied(oldText, hunks)).toBe(newText)
      hunks.forEach((hunk, i) => {
        expect(hunk.from).toBeLessThanOrEqual(hunk.to)
        if (i > 0) expect(hunks[i - 1].to).toBeLessThanOrEqual(hunk.from)
      })
    }
  })
})
