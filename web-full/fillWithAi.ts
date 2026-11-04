// "Fill with AI" from the chip, the context menu and the host's bar: what it should fill is the field under the
// caret, or else the selected text and its paragraph, which the host hands to the chat.
import type { EditorView } from '@tiptap/pm/view'
import type { EditorToHost } from './embed-protocol'
import { caretIds } from './fieldReporting'

const LABELS: Record<string, string> = {
  de: 'Mit KI ausfüllen',
  fr: "Remplir avec l'IA",
  bg: 'Попълни с ИИ',
}

// An empty line says nothing itself, so the text leading up to it does.
const TEXT_BEFORE_EMPTY_LINE = 200

/** the entry's name in the embed's UI language */
export const fillWithAiLabel = (lang: string) => LABELS[lang] ?? 'Fill with AI'

/** what the caret or selection asks the host to fill */
export const fillWithAiMessage = (view: EditorView): EditorToHost => {
  const { selection, doc } = view.state
  return {
    type: 'vound:fillWithAi',
    ids: caretIds(view),
    selected: doc.textBetween(selection.from, selection.to, ' '),
    paragraph: selection.$from.parent.textContent || doc.textBetween(Math.max(0, selection.from - TEXT_BEFORE_EMPTY_LINE), selection.from, ' ').trim(),
  }
}
