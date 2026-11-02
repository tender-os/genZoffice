import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { AI_MENU_KEYS, menuShows, REVIEW_INSERT_KEYS, REVIEW_MENU_KEYS } from '../reviewMenu'

const contextMenu = readFileSync(resolve(__dirname, '../../apps/docs/src/renderer/components/ContextMenu.tsx'), 'utf8')

describe('menuShows', () => {
  it('leaves the AI entries out of the embed, in review mode too', () => {
    for (const key of AI_MENU_KEYS) {
      expect(menuShows(key, true, false)).toBe(false)
      expect(menuShows(key, true, true)).toBe(false)
      expect(menuShows(key, false, false)).toBe(true)
    }
  })

  it('leaves formatted Paste, the field commands and New Comment out of review mode only', () => {
    for (const key of REVIEW_MENU_KEYS) {
      expect(menuShows(key, true, true)).toBe(false)
      expect(menuShows(key, true, false)).toBe(true)
    }
  })

  it('leaves the commands that change a form\'s structure out of review mode only', () => {
    for (const key of ['ribbonTableDeleteMenu', 'ribbonMergeCells', 'ribbonAutoFit', 'appRestartNumbering', ...REVIEW_INSERT_KEYS]) {
      expect(menuShows(key, true, true)).toBe(false)
      expect(menuShows(key, true, false)).toBe(true)
    }
  })

  it('keeps plain paste, formatting, clipboard, hyperlinks and adding rows in review mode', () => {
    for (const key of ['appPastePlain', 'appCut', 'appCopy', 'appFontMenu', 'appParagraphMenu', 'appHyperlinkMenu', 'ribbonInsert', 'ribbonInsertAbove', 'ribbonInsertBelow']) {
      expect(menuShows(key, true, true)).toBe(true)
    }
  })

  it('finds every entry once in the context menu, as the transform needs', () => {
    for (const key of [...AI_MENU_KEYS, ...REVIEW_MENU_KEYS]) {
      expect(contextMenu.split(`item(t('${key}'),`)).toHaveLength(2)
    }
  })

  it('finds the Insert entries that add columns or cells, as the transform needs', () => {
    for (const needle of ["['ribbonInsertLeft', addColumnBefore],", "['ribbonInsertRight', addColumnAfter],", "onClick={run(() => onTableDialog('insertCells'))}"]) {
      expect(contextMenu.split(needle)).toHaveLength(2)
    }
  })
})
