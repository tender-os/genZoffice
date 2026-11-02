// The context-menu entries the embed leaves out (the reviewMenu transform in vite.config.ts asks menuShows
// before each one): the AI entries and hiding text from the AI everywhere in the embed, since the AI is
// hidden; in review mode also formatted Paste (Paste as plain text stays), the field commands and New
// Comment, whose results the host does not follow, and every command that changes a form's structure:
// changes to tender documents can exclude a bid, and the host's change list compares text only.
// Formatting, clipboard, hyperlinks and adding table rows stay.

/** the entries' string keys in apps/docs ContextMenu.tsx */
export const AI_MENU_KEYS = ['appSynonyms', 'appTranslate', 'redactMenuLabel', 'redactShowLabel'] as const
export const REVIEW_MENU_KEYS = [
  'appPaste',
  'appUpdateField',
  'appToggleFieldCodes',
  'appEditField',
  'appNewComment',
  'ribbonTableDeleteMenu',
  'ribbonMergeCells',
  'ribbonSplitCellsDialog',
  'ribbonSplitTable',
  'ribbonDistribute',
  'ribbonAutoFit',
  'ribbonTableProperties',
  'appRestartNumbering',
  'appContinueNumbering',
  'appSetNumberingValue',
] as const
/** the Insert submenu's entries that add columns or cells; adding rows stays */
export const REVIEW_INSERT_KEYS = ['ribbonInsertLeft', 'ribbonInsertRight', 'ribbonInsertCells'] as const

const aiKeys: readonly string[] = AI_MENU_KEYS
const reviewKeys: readonly string[] = [...REVIEW_MENU_KEYS, ...REVIEW_INSERT_KEYS]

export function menuShows(key: string, embedded: boolean, reviewMode: boolean): boolean {
  if (embedded && aiKeys.includes(key)) return false
  return !(reviewMode && reviewKeys.includes(key))
}
