import { escapeXmlAttr } from './xml-utils'

/** the Latin face (w:rFonts w:ascii, w:hAnsi) and the size in half-points (w:sz, w:szCs) of a run */
/** a size of null takes the run's own size away */
export type FaceAndSize = { fontAscii?: string; sizeHalfPoints?: number | null }

// CT_RPr children that come after w:sz and w:szCs
const AFTER_SIZE = [
  'highlight',
  'u',
  'effect',
  'bdr',
  'shd',
  'fitText',
  'vertAlign',
  'rtl',
  'cs',
  'em',
  'lang',
  'eastAsianLayout',
  'specVanish',
  'oMath',
]

const RPR_RE = /<w:rPr>[\s\S]*?<\/w:rPr>|<w:rPr\/>/

const attrOf = (element: string | undefined, name: string) =>
  element ? new RegExp(`\\s${name}="([^"]*)"`).exec(element)?.[1] : undefined

const elementOf = (rPr: string, tag: string) => new RegExp(`<w:${tag}\\b[^>]*/>`).exec(rPr)?.[0]

/** the face and size a w:r's own run properties declare; `themed` when its Latin face comes from the theme */
export function declaredFaceAndSize(runXml: string) {
  const rPr = RPR_RE.exec(runXml)?.[0] ?? ''
  const fonts = elementOf(rPr, 'rFonts')
  const size = attrOf(elementOf(rPr, 'sz'), 'w:val')
  return {
    fontAscii: attrOf(fonts, 'w:ascii'),
    sizeHalfPoints: size === undefined ? undefined : Number(size),
    themed: attrOf(fonts, 'w:asciiTheme') !== undefined,
  }
}

/** an rPr child put where CT_RPr orders it: before the first of `before` it holds, else at the end */
export function placed(inner: string, child: string, before: readonly string[]) {
  const at = before
    .map((tag) => inner.search(new RegExp(`<w:${tag}[\\s/>]`)))
    .filter((index) => index !== -1)
    .sort((a, b) => a - b)[0]
  return at === undefined ? inner + child : inner.slice(0, at) + child + inner.slice(at)
}

function withFace(inner: string, face: string) {
  const fonts = elementOf(inner, 'rFonts')
  const value = escapeXmlAttr(face)
  if (!fonts) {
    const style = elementOf(inner, 'rStyle')
    const child = `<w:rFonts w:ascii="${value}" w:hAnsi="${value}"/>`
    return style ? inner.replace(style, style + child) : child + inner
  }
  const kept = fonts
    .replace(/\s(?:w:ascii|w:hAnsi|w:asciiTheme|w:hAnsiTheme)="[^"]*"/g, '')
    .replace(/\s*\/>$/, '')
  return inner.replace(fonts, `${kept} w:ascii="${value}" w:hAnsi="${value}"/>`)
}

function withSize(inner: string, halfPoints: number | null) {
  const unsized = inner.replace(/<w:szCs\b[^>]*\/>/, '').replace(/<w:sz\b[^>]*\/>/, '')
  if (halfPoints === null) return unsized
  return placed(unsized, `<w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/>`, AFTER_SIZE)
}

/** a w:r with the face and size set in its run properties; every other property stays */
export function withFaceAndSize(runXml: string, style: FaceAndSize): string {
  const rPr = RPR_RE.exec(runXml)?.[0]
  let inner = rPr && rPr !== '<w:rPr/>' ? rPr.slice('<w:rPr>'.length, -'</w:rPr>'.length) : ''
  if (style.fontAscii !== undefined) inner = withFace(inner, style.fontAscii)
  if (style.sizeHalfPoints !== undefined) inner = withSize(inner, style.sizeHalfPoints)
  const next = `<w:rPr>${inner}</w:rPr>`
  if (rPr) return runXml.replace(rPr, next)
  const open = /^<w:r(?:\s[^>]*)?>/.exec(runXml)?.[0] ?? '<w:r>'
  return open + next + runXml.slice(open.length)
}
