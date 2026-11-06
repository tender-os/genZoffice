// A reveal names text in the document the host just opened, and the editor loads that document after
// vound:open returns. So the editor tells the host when the opened document is in (vound:opened), and a
// reveal that arrives before then waits for it; a newer reveal or open answers the waiting one with
// found: false.
import type { EditorView } from '@tiptap/pm/view'
import type { EditorToHost, HostToEditor } from './embed-protocol'
import { revealPassage } from './reveal'

type RevealRequest = Extract<HostToEditor, { type: 'vound:reveal' }>

/** Posts vound:opened once per open and answers every reveal exactly once, after the document it names has arrived. */
export function createRevealQueue(post: (msg: EditorToHost) => void) {
  let loading = false
  let waiting: RevealRequest | null = null
  const answer = (requestId: string, found: boolean) => post({ type: 'vound:revealResult', requestId, found })
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
      post({ type: 'vound:opened' })
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
