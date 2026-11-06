// A reveal names text in the document the host just opened, and the editor loads that document after
// vound:open returns. So a reveal that arrives while the opened document is not in the editor yet waits
// for it; a newer reveal or open answers the waiting one with found: false.
import type { EditorView } from '@tiptap/pm/view'
import type { HostToEditor } from './embed-protocol'
import { revealPassage } from './reveal'

type RevealRequest = Extract<HostToEditor, { type: 'vound:reveal' }>

/** Answers every reveal exactly once through `answer`, after the document it names has arrived. */
export function createRevealQueue(answer: (requestId: string, found: boolean) => void) {
  let loading = false
  let waiting: RevealRequest | null = null
  const dropWaiting = () => {
    if (waiting) answer(waiting.requestId, false)
    waiting = null
  }
  return {
    /** vound:open: the next document change brings the document the host opened */
    opening() {
      dropWaiting()
      loading = true
    },
    documentChanged(view: EditorView) {
      if (!loading) return
      loading = false
      if (waiting) answer(waiting.requestId, revealPassage(view, waiting))
      waiting = null
    },
    reveal(view: EditorView | null, request: RevealRequest) {
      dropWaiting()
      if (view && !loading) answer(request.requestId, revealPassage(view, request))
      else waiting = request
    },
  }
}
