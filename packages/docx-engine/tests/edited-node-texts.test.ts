import { describe, expect, test } from 'vitest'
import { editedNodeTexts } from '../src/edited-node-texts'

describe('editedNodeTexts (Vound fork)', () => {
  test('text typed at the end of a run stays in that run, not in a later field', () => {
    expect(
      editedNodeTexts('Firma: Muster GmbH & Co. ja', ['Firma: ', 'Muster GmbH', ' ', 'ja']),
    ).toEqual(['Firma: ', 'Muster GmbH & Co.', ' ', 'ja'])
  })
  test('a change across two runs goes into the first, and the second keeps its tail', () => {
    expect(editedNodeTexts('Ort: Köln Mitte', ['Ort: Ber', 'lin Mitte'])).toEqual([
      'Ort: Köl',
      'n Mitte',
    ])
  })
  test('an unchanged text gives nothing to write', () => {
    expect(editedNodeTexts('ab', ['a', 'b'])).toBeNull()
  })
})
