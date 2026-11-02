// The context-menu entries the embed leaves out (the reviewMenu transform in vite.config.ts asks menuShows
// before each one): the AI entries and hiding text from the AI everywhere in the embed, since the AI is
// hidden; in review mode also formatted Paste (Paste as plain text stays), the field commands and New
// Comment, whose results the host does not follow. Formatting, clipboard, hyperlinks and tables stay.

/** the entries' string keys in apps/docs ContextMenu.tsx */
export const AI_MENU_KEYS = ['appSynonyms', 'appTranslate', 'redactMenuLabel', 'redactShowLabel'] as const
export const REVIEW_MENU_KEYS = ['appPaste', 'appUpdateField', 'appToggleFieldCodes', 'appEditField', 'appNewComment'] as const

const aiKeys: readonly string[] = AI_MENU_KEYS
const reviewKeys: readonly string[] = REVIEW_MENU_KEYS

export function menuShows(key: string, embedded: boolean, reviewMode: boolean): boolean {
  if (embedded && aiKeys.includes(key)) return false
  return !(reviewMode && reviewKeys.includes(key))
}
