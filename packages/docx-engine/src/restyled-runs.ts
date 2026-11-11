import { formTextResultSpans } from './form-text-anchor'
import { decodeEntities } from './parse-xml-text'
import { declaredFaceAndSize, withFaceAndSize, type FaceAndSize } from './run-face-size'
import { wtTextOf, type WtMappedRun } from './text-patch'
import type { Run } from './types'
import { escapeXmlText } from './xml-utils'

/** a run of the editor's model, with the text the file holds in w:t and its face and size */
/** a run of the editor's model, with the text the file holds in w:t and its face and size; a size of null, or none
 *  on a run whose properties had one (rawRPr), is a size the editor took away */
export type RestyledRun = WtMappedRun &
  Pick<Run, 'fontAscii' | 'themeRFonts' | 'rawRPr'> & { sizeHalfPoints?: number | null }

type FileRun = { start: number; end: number; xml: string; text: string; simple: boolean }

const RUN_RE = /<w:r(?=[\s>])[\s\S]*?<\/w:r>/g
const T_RE = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:t(?:\s[^>]*)?\/>/g
const RPR_RE = /<w:rPr>[\s\S]*?<\/w:rPr>|<w:rPr\/>/

/** the w:r of the paragraph that hold text, with that text; `simple` when it holds nothing else */
function fileRuns(xml: string): FileRun[] {
  return [...xml.matchAll(RUN_RE)].flatMap((match) => {
    const texts = [...match[0].matchAll(T_RE)]
    if (texts.length === 0) return []
    const text = texts.map((t) => decodeEntities(t[1] ?? '')).join('')
    const rest = match[0]
      .replace(/^<w:r(?:\s[^>]*)?>/, '')
      .replace(/<\/w:r>$/, '')
      .replace(RPR_RE, '')
      .replace(T_RE, '')
    return [
      {
        start: match.index,
        end: match.index + match[0].length,
        xml: match[0],
        text,
        simple: texts.length === 1 && rest.trim() === '',
      },
    ]
  })
}

const modelText = (run: RestyledRun) => wtTextOf(run).replace(/[\t\n]/g, '')

/** the size the model asks of a run: its own, none when it was taken away, undefined when it says nothing */
function sizeAsked(run: RestyledRun) {
  if (run.sizeHalfPoints !== undefined) return run.sizeHalfPoints
  return run.rawRPr?.includes('<w:sz ') ? null : undefined
}

/** the face and size the model asks of a file run that does not already show them, or null */
function wantedOf(run: RestyledRun, file: FileRun): FaceAndSize | null {
  const declared = declaredFaceAndSize(file.xml)
  const face =
    declared.themed && declared.fontAscii === undefined
      ? run.themeRFonts?.fontAscii
      : declared.fontAscii
  const newFace = run.fontAscii !== undefined && run.fontAscii !== face
  const size = sizeAsked(run)
  const newSize = size !== undefined && size !== (declared.sizeHalfPoints ?? null)
  if (!newFace && !newSize) return null
  return {
    ...(newFace ? { fontAscii: run.fontAscii } : {}),
    ...(newSize ? { sizeHalfPoints: size } : {}),
  }
}

/** the model run behind each character of `text`, or null when the model's text is not found in it once */
function ownersOf(text: string, runs: ReadonlyArray<RestyledRun>) {
  const owners = runs.flatMap((run, index) => [...modelText(run)].map(() => index))
  const joined = runs.map(modelText).join('')
  const at = text.indexOf(joined)
  if (at === -1 || text.indexOf(joined, at + 1) !== -1) return null
  return (offset: number) => owners[offset - at]
}

/** consecutive characters of a file run that belong to one model run */
function groupsOf(file: FileRun, from: number, owner: (offset: number) => number | undefined) {
  const groups: { index: number | undefined; text: string }[] = []
  ;[...file.text].forEach((char, i) => {
    const index = owner(from + i)
    const last = groups.at(-1)
    if (last && last.index === index) last.text += char
    else groups.push({ index, text: char })
  })
  return groups
}

/** Vound fork: the face and size the editor gave runs written into the paragraph's w:r, where a text edit keeps
 *  the original bytes (form fields, content controls). A w:r that already shows them keeps its bytes; one whose
 *  text took two styles is split, when it holds nothing but its text. A model whose text the file does not hold
 *  once leaves the paragraph alone. */
export function withRestyledRuns(xml: string, runs: ReadonlyArray<RestyledRun>): string {
  const files = fileRuns(xml)
  const owner = ownersOf(files.map((file) => file.text).join(''), runs)
  if (!owner) return xml
  let out = xml
  let offset = files.reduce((sum, file) => sum + file.text.length, 0)
  for (const file of [...files].reverse()) {
    offset -= file.text.length
    const groups = groupsOf(file, offset, owner)
    const wanted = groups.map((group) =>
      group.index === undefined ? null : wantedOf(runs[group.index], file),
    )
    if (wanted.every((style) => style === null)) continue
    out = out.slice(0, file.start) + rebuilt(file, groups, wanted) + out.slice(file.end)
  }
  return out
}

/** the file run with each group's face and size: split into one run per group when it holds only its text */
function rebuilt(file: FileRun, groups: { text: string }[], wanted: (FaceAndSize | null)[]) {
  if (groups.length === 1) return withFaceAndSize(file.xml, wanted[0] ?? {})
  if (!file.simple) return file.xml
  return groups
    .map((group, i) => withFaceAndSize(textRun(file.xml, group.text), wanted[i] ?? {}))
    .join('')
}

/** a simple w:r with another text */
const textRun = (runXml: string, text: string) =>
  runXml.replace(T_RE, `<w:t xml:space="preserve">${escapeXmlText(text)}</w:t>`)

/** Vound fork: the face and size on every run of the paragraph's legacy text form field results that does not show
 *  them already */
export function withResultFaceAndSize(xml: string, style: FaceAndSize): string {
  const spans = formTextResultSpans(xml)
  const inResult = (file: FileRun) =>
    spans.some((span) => file.start > span.from && file.start < span.to)
  const model = { text: '', ...style }
  return fileRuns(xml)
    .filter(inResult)
    .reverse()
    .reduce((out, file) => {
      const wanted = wantedOf(model, file)
      return wanted
        ? out.slice(0, file.start) + withFaceAndSize(file.xml, wanted) + out.slice(file.end)
        : out
    }, xml)
}
