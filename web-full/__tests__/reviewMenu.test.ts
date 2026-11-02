import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { AI_MENU_KEYS, menuShows, REVIEW_MENU_KEYS } from '../reviewMenu'

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

  it('keeps plain paste, formatting, clipboard and hyperlinks in review mode', () => {
    for (const key of ['appPastePlain', 'appCut', 'appCopy', 'appFontMenu', 'appParagraphMenu', 'appHyperlinkMenu']) {
      expect(menuShows(key, true, true)).toBe(true)
    }
  })

  it('finds every entry once in the context menu, as the transform needs', () => {
    for (const key of [...AI_MENU_KEYS, ...REVIEW_MENU_KEYS]) {
      expect(contextMenu.split(`item(t('${key}'),`)).toHaveLength(2)
    }
  })
})
