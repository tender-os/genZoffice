import { describe, expect, it } from 'vitest'
import { blockedInReview, keepsNativeAction, type ShortcutKeys } from '../reviewShortcuts'

const keys = (init: Partial<ShortcutKeys>): ShortcutKeys => ({ key: '', code: '', metaKey: false, ctrlKey: false, shiftKey: false, altKey: false, ...init })
const letter = (key: string, init: Partial<ShortcutKeys> = {}) => keys({ key, code: `Key${key.toUpperCase()}`, ...init })
const ctrl = { ctrlKey: true }
const cmd = { metaKey: true }

describe('blockedInReview', () => {
  it('blocks the document tools the host does not follow and File › Open', () => {
    const blocked = [
      // fields, track changes, new comment, footnote and endnote, date and time fields, AI proofread
      keys({ key: 'F9', code: 'F9' }), keys({ altKey: true, key: 'F9', code: 'F9' }),
      letter('e', { ...ctrl, shiftKey: true }), letter('a', { ...ctrl, altKey: true }), letter('f', { ...ctrl, altKey: true }),
      letter('d', { ...ctrl, altKey: true }), letter('d', { altKey: true, shiftKey: true }), letter('t', { altKey: true, shiftKey: true }),
      keys({ key: 'F7', code: 'F7' }),
      // File › Open: another file would replace the reviewed document
      letter('o', ctrl),
    ]
    expect(blocked.filter((event) => !blockedInReview(event, false))).toEqual([])
  })

  it('keeps the formatting shortcuts, as Word does', () => {
    const formatting = [
      // bold, italic, underline; alignment left, center, right, justify
      ...['b', 'i', 'u', 'l', 'e', 'r', 'j'].map((key) => letter(key, ctrl)),
      // font dialog, hyperlink, paragraph dialog
      letter('d', ctrl), letter('k', ctrl), letter('µ', { ...ctrl, altKey: true, code: 'KeyM' }),
      // line spacing 1 / 2 / 1.5
      ...['1', '2', '5'].map((digit) => keys({ ...ctrl, key: digit, code: `Digit${digit}` })),
      // paragraph styles Normal, Heading 1-3
      ...['0', '1', '2', '3'].map((digit) => keys({ ...ctrl, altKey: true, key: digit, code: `Digit${digit}` })),
      // font size grow / shrink, one-point nudge
      keys({ ...ctrl, shiftKey: true, key: '>', code: 'Period' }), keys({ ...ctrl, shiftKey: true, key: '<', code: 'Comma' }),
      keys({ ...ctrl, key: ']', code: 'BracketRight' }), keys({ ...ctrl, key: '[', code: 'BracketLeft' }),
      // superscript / subscript
      keys({ ...ctrl, key: '.', code: 'Period' }), keys({ ...ctrl, key: ',', code: 'Comma' }), keys({ ...ctrl, shiftKey: true, key: '+', code: 'Equal' }),
      // hanging indent, indent / outdent, clear character and paragraph formatting, change case
      letter('t', ctrl), letter('t', { ...ctrl, shiftKey: true }), letter('m', ctrl), letter('m', { ...ctrl, shiftKey: true }),
      keys({ ...ctrl, key: ' ', code: 'Space' }), letter('q', ctrl), keys({ shiftKey: true, key: 'F3', code: 'F3' }),
    ]
    expect(formatting.filter((event) => blockedInReview(event, false))).toEqual([])
    expect(formatting.map((event) => ({ ...event, ctrlKey: false, metaKey: true })).filter((event) => blockedInReview(event, true))).toEqual([])
  })

  it('keeps undo, redo, clipboard, select-all, save, find, zoom, navigation and typing', () => {
    const allowed = [
      letter('z', ctrl), letter('z', { ...ctrl, shiftKey: true }), letter('y', ctrl),
      letter('c', ctrl), letter('x', ctrl), letter('v', ctrl), letter('a', ctrl), letter('s', ctrl), letter('f', ctrl),
      keys({ ...ctrl, key: '=', code: 'Equal' }), keys({ ...ctrl, key: '-', code: 'Minus' }),
      keys({ key: 'ArrowLeft', code: 'ArrowLeft' }), keys({ ...ctrl, key: 'ArrowLeft', code: 'ArrowLeft' }),
      keys({ shiftKey: true, key: 'ArrowRight', code: 'ArrowRight' }), keys({ key: 'Home', code: 'Home' }), keys({ ...ctrl, key: 'End', code: 'End' }),
      keys({ ...ctrl, key: 'Backspace', code: 'Backspace' }), keys({ key: 'Tab', code: 'Tab' }), keys({ key: 'Enter', code: 'Enter' }),
      letter('e'), letter('E', { shiftKey: true }), keys({ key: '8', code: 'Digit8' }),
      // AltGr on Windows arrives as Ctrl+Alt: @ and € on a German layout
      letter('@', { ...ctrl, altKey: true, code: 'KeyQ' }), letter('€', { ...ctrl, altKey: true, code: 'KeyE' }),
    ]
    expect(allowed.filter((event) => blockedInReview(event, false))).toEqual([])
  })

  it('uses Command on the Mac and keeps its native Control navigation', () => {
    expect(blockedInReview(letter('b', cmd), true)).toBe(false)
    expect(blockedInReview(letter('o', cmd), true)).toBe(true)
    // the endnote shortcut is ⌥⌘E on the Mac, Ctrl+Alt+D elsewhere
    expect(blockedInReview(letter('´', { ...cmd, altKey: true, code: 'KeyE' }), true)).toBe(true)
    expect(blockedInReview(letter('∂', { ...cmd, altKey: true, code: 'KeyD' }), true)).toBe(false)
    // the app also aligns on ⌃E and bolds on ⌃B; the blocker stops that, the native move stays
    expect(blockedInReview(letter('e', ctrl), true)).toBe(true)
    expect(blockedInReview(letter('b', ctrl), true)).toBe(true)
    expect(keepsNativeAction(letter('e', ctrl), true)).toBe(true)
    expect(keepsNativeAction(letter('e', cmd), true)).toBe(false)
    expect(keepsNativeAction(letter('e', ctrl), false)).toBe(false)
  })
})
