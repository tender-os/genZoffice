// The keyboard shortcuts review mode blocks: the document tools whose results the host does not follow
// (track changes, comments, footnotes and endnotes, fields, AI proofread) and File › Open, whose file would
// replace the reviewed document. Formatting passes, as in Word: bold, italic, underline, alignment, font
// size, the font, paragraph and hyperlink dialogs, spacing, styles, indents and clearing formatting.
// Matching mirrors App.tsx: Mod is Command or Control on every platform; letters it reads from e.key,
// digits and punctuation from e.code.
export type ShortcutKeys = Pick<KeyboardEvent, 'key' | 'code' | 'metaKey' | 'ctrlKey' | 'shiftKey' | 'altKey'>

/** `mod`: 'any' is Command or Control, 'none' neither. An unset
 *  modifier may be either way. `key` is the lower-case letter, `code` the physical key. */
type Combo = { mod?: 'any' | 'none'; shift?: boolean; alt?: boolean; key?: string; code?: string }

const codes = (base: Combo, ...list: string[]): Combo[] => list.map((code) => ({ ...base, code }))

const combos = (mac: boolean): Combo[] => [
  // File › Open: App.tsx opens on Mod+o with any Shift or Alt
  { mod: 'any', key: 'o' },
  // track changes ⇧⌘E
  { mod: 'any', alt: false, shift: true, code: 'KeyE' },
  // new comment ⌥⌘A; footnote ⌥⌘F; endnote ⌥⌘E on the Mac, Ctrl+Alt+D elsewhere (Ctrl+Alt+E is AltGr+E there)
  ...codes({ mod: 'any', alt: true }, 'KeyA', 'KeyF', mac ? 'KeyE' : 'KeyD'),
  // date and time fields ⌥⇧D / ⌥⇧T
  ...codes({ mod: 'none', alt: true, shift: true }, 'KeyD', 'KeyT'),
  // update fields F9 and field codes ⌥F9; AI proofread F7
  { code: 'F9' },
  { shift: false, code: 'F7' },
]

const holds = (expected: boolean | undefined, actual: boolean) => expected === undefined || expected === actual
const modHeld = { any: (e: ShortcutKeys) => e.metaKey || e.ctrlKey, none: (e: ShortcutKeys) => !e.metaKey && !e.ctrlKey }

function matches(combo: Combo, event: ShortcutKeys) {
  if (combo.mod && !modHeld[combo.mod](event)) return false
  if (!holds(combo.shift, event.shiftKey) || !holds(combo.alt, event.altKey)) return false
  if (combo.key !== undefined && combo.key !== event.key.toLowerCase()) return false
  return combo.code === undefined || combo.code === event.code
}

export const isMacPlatform = () => typeof navigator !== 'undefined' && /Mac|iP(hone|[oa]d)/.test(navigator.platform)

/** the Mac's Control+letter: App.tsx reads it as Mod+letter, the system as text navigation (⌃E end of line) */
const macControlLetter = (event: ShortcutKeys, mac: boolean) => mac && event.ctrlKey && !event.metaKey && /^Key[A-Z]$/.test(event.code)

export function blockedInReview(event: ShortcutKeys, mac: boolean): boolean {
  return macControlLetter(event, mac) || combos(mac).some((combo) => matches(combo, event))
}

/** On the Mac, Control+letter is the system's text navigation (⌃E end of line, ⌃D delete forward):
 *  the app and the editor must not see it, but the browser's own action stays. */
export function keepsNativeAction(event: ShortcutKeys, mac: boolean): boolean {
  return mac && event.ctrlKey && !event.metaKey
}
