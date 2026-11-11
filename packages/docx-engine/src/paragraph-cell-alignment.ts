import { placed } from './run-face-size'
import type { ParaAlign, TableCell } from './types'

const JC = {
  left: 'left',
  center: 'center',
  right: 'right',
  justify: 'both',
  distribute: 'distribute',
} satisfies Record<ParaAlign, string>

// CT_PPr children after w:jc, and CT_TcPr children after w:vAlign
const AFTER_JC = [
  'textDirection',
  'textAlignment',
  'textboxTightWrap',
  'outlineLvl',
  'divId',
  'cnfStyle',
  'rPr',
  'sectPr',
  'pPrChange',
]
const AFTER_VALIGN = ['hideMark', 'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange']

/** the element's properties child (w:pPr, w:tcPr) set to `child`, in schema order, after removing `tag` */
function withProperty(
  xml: string,
  element: 'p' | 'tc',
  properties: 'pPr' | 'tcPr',
  tag: string,
  child: string,
  after: readonly string[],
) {
  const open = new RegExp(`^<w:${element}(?:\\s[^>]*)?>`).exec(xml)?.[0]
  if (!open) return xml
  const rest = xml.slice(open.length)
  const found = new RegExp(
    `^<w:${properties}>[\\s\\S]*?</w:${properties}>|^<w:${properties}/>`,
  ).exec(rest)?.[0]
  const inner =
    found && found !== `<w:${properties}/>`
      ? found.slice(properties.length + 4, -(properties.length + 5))
      : ''
  const without = inner.replace(new RegExp(`<w:${tag}\\b[^>]*/>`), '')
  const next = `<w:${properties}>${placed(without, child, after)}</w:${properties}>`
  return open + next + rest.slice(found?.length ?? 0)
}

/** Vound fork: a paragraph with its alignment (w:jc) set; every other property stays */
export const withParagraphAlignment = (paraXml: string, align: ParaAlign) =>
  withProperty(paraXml, 'p', 'pPr', 'jc', `<w:jc w:val="${JC[align]}"/>`, AFTER_JC)

/** Vound fork: a table cell with its vertical alignment (w:vAlign) set; every other property stays */
export const withCellVerticalAlignment = (
  cellXml: string,
  vAlign: NonNullable<TableCell['vAlign']>,
) => withProperty(cellXml, 'tc', 'tcPr', 'vAlign', `<w:vAlign w:val="${vAlign}"/>`, AFTER_VALIGN)
