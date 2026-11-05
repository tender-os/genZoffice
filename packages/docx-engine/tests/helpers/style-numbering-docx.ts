import { buildDocx } from './build-docx'

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const STYLES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}>` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
  '<w:pPr><w:ind w:left="1000"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:ind w:left="499" w:hanging="357"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Bar"><w:name w:val="Bar"/><w:basedOn w:val="Heading2"/>' +
  '<w:pPr><w:ind w:right="-144"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Plain"><w:name w:val="Plain"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>' +
  '</w:styles>'

const NUMBERING_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${NS}>` +
  '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
  '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>' +
  '</w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0">' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
  '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="9999" w:hanging="360"/></w:pPr></w:lvl>' +
  '</w:lvlOverride></w:num>' +
  '</w:numbering>'

/** Heading styles that carry numbering whose level indent (9999) differs from the style's (499). */
export function buildStyleNumberingDocx(bodyXml: string): Promise<Uint8Array> {
  return buildDocx({ bodyXml, stylesXml: STYLES_XML, numberingXml: NUMBERING_XML })
}

export const DIRECT_NUMPR = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
