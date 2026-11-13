import * as xml from '../xml.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Package } from '../opc/package.ts'

export interface Fonts {
  heading: string
  body: string
}

export interface CreateOptions {
  title?: string
  page?: 'letter' | 'a4'
  size?: number
  fonts?: Partial<Fonts>
}

interface Page {
  width: number
  height: number
}

interface Resolved {
  page: Page
  fonts: Fonts
  size: number
  title: unknown
}

interface HeadingStyle {
  id: string
  name: string
  level: number
  size: number
  color: string
  before: number
}

interface ListLevel {
  ilvl: number
  format: string
  text: string
  font: string | null
}

const W = xml.ns.w
const R = xml.ns.r

const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'
const TYPES = {
  document: WML + 'document.main+xml',
  styles: WML + 'styles+xml',
  numbering: WML + 'numbering+xml',
  settings: WML + 'settings+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
}

export const PAGES: Record<string, Page> = {
  letter: { width: 12240, height: 15840 },
  a4: { width: 11906, height: 16838 }
}

const BULLETS = ['\u2022', 'o', '\u25AA']
const BULLET_FONTS = [null, 'Courier New', 'Wingdings']
const NUMBER_FORMATS = ['decimal', 'lowerLetter', 'lowerRoman']

function escape(value: unknown): string {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function fontsOf(input: unknown): Fonts {
  const fonts: Record<string, unknown> = { heading: 'Calibri Light', body: 'Calibri' }
  if (isRecord(input)) {
    Object.assign(fonts, input)
  }

  for (const [key, value] of Object.entries(fonts)) {
    if (!['heading', 'body'].includes(key)) {
      throw new DocsError('INVALID_OPTIONS', `unknown font "${key}"; use heading or body`)
    }

    if (typeof value !== 'string' || !value.trim()) {
      throw new DocsError('INVALID_OPTIONS', `font ${key} must be a typeface name`)
    }
  }

  return { heading: String(fonts.heading), body: String(fonts.body) }
}

function options(opts: Record<string, unknown>): Resolved {
  const page = PAGES[String(opts.page || 'letter').toLowerCase()]
  if (!page) {
    throw new DocsError('INVALID_OPTIONS', 'page must be "letter" or "a4"')
  }

  const fonts = fontsOf(opts.fonts)
  const size = opts.size === undefined ? 11 : opts.size
  if (!(typeof size === 'number' && size >= 6 && size <= 72)) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `size is the body text size in points, from 6 to 72, got ${String(size)}`
    )
  }

  return { page, fonts, size, title: opts.title }
}

function fontsXml(face: string): string {
  const f = escape(face)
  return `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/>`
}

function heading(style: HeadingStyle): string {
  const { id, name, level, size, color, before } = style
  return `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${before}" w:after="0"/><w:outlineLvl w:val="${level}"/></w:pPr><w:rPr>{heading}<w:color w:val="${color}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`
}

function stylesXml(fonts: Fonts, size: number): string {
  const half = Math.round(size * 2)
  const body = fontsXml(fonts.body)
  const head = fontsXml(fonts.heading)
  const styles = [
    `<w:docDefaults><w:rPrDefault><w:rPr>${body}<w:sz w:val="${half}"/><w:szCs w:val="${half}"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`,
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>',
    '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>',
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>',
    '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>',
    `<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr>${head}<w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>`,
    `<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="11"/><w:qFormat/><w:rPr><w:color w:val="5A5A5A"/><w:spacing w:val="15"/><w:sz w:val="${half + 6}"/><w:szCs w:val="${half + 6}"/></w:rPr></w:style>`,
    heading({
      id: 'Heading1',
      name: 'heading 1',
      level: 0,
      size: 32,
      color: '2F5496',
      before: 240
    }),
    heading({ id: 'Heading2', name: 'heading 2', level: 1, size: 26, color: '2F5496', before: 40 }),
    heading({ id: 'Heading3', name: 'heading 3', level: 2, size: 24, color: '1F3763', before: 40 }),
    '<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:contextualSpacing/></w:pPr></w:style>',
    '<w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr><w:contextualSpacing/></w:pPr></w:style>',
    '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="29"/><w:qFormat/><w:pPr><w:spacing w:before="200"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr></w:style>',
    '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>'
  ]

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="${W}">${styles.join('').replace(/\{heading\}/g, head)}</w:styles>`
}

function levelXml(level: ListLevel): string {
  const font = level.font
    ? `<w:rPr><w:rFonts w:ascii="${level.font}" w:hAnsi="${level.font}" w:hint="default"/></w:rPr>`
    : ''
  return `<w:lvl w:ilvl="${level.ilvl}"><w:start w:val="1"/><w:numFmt w:val="${level.format}"/><w:lvlText w:val="${level.text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${360 + 360 * level.ilvl}" w:hanging="360"/></w:pPr>${font}</w:lvl>`
}

function numberingXml(): string {
  const nine = Array.from({ length: 9 }, (_, i) => i)
  const bulletLevels = nine
    .map((i) =>
      levelXml({ ilvl: i, format: 'bullet', text: BULLETS[i % 3], font: BULLET_FONTS[i % 3] })
    )
    .join('')
  const numberLevels = nine
    .map((i) =>
      levelXml({ ilvl: i, format: NUMBER_FORMATS[i % 3], text: `%${i + 1}.`, font: null })
    )
    .join('')

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${bulletLevels}</w:abstractNum><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${numberLevels}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
}

function documentXml(page: Page): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${xml.ns.wp}" xmlns:a="${xml.ns.a}" xmlns:pic="${xml.ns.pic}"><w:body><w:p/><w:sectPr><w:pgSz w:w="${page.width}" w:h="${page.height}"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr></w:body></w:document>`
}

function addProperties(pkg: Package, title: unknown): void {
  const heading = title ? `<dc:title>${escape(title)}</dc:title>` : ''
  pkg.add(
    '/docProps/core.xml',
    xml.parse(
      `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">${heading}</cp:coreProperties>`
    ),
    TYPES.core
  )
  pkg.add(
    '/docProps/app.xml',
    xml.parse(
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>QVAC</Application></Properties>'
    ),
    TYPES.app
  )
  const rels = pkg.rels('/')
  rels.add(REL_BASE + 'officeDocument', '/word/document.xml')
  rels.add(
    'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
    '/docProps/core.xml'
  )
  rels.add(REL_BASE + 'extended-properties', '/docProps/app.xml')
}

export function build(pkg: Package, input: unknown = {}): void {
  const { page, fonts, size, title } = options(isRecord(input) ? input : {})
  pkg.add('/word/document.xml', xml.parse(documentXml(page)), TYPES.document)
  pkg.add('/word/styles.xml', xml.parse(stylesXml(fonts, size)), TYPES.styles)
  pkg.add('/word/numbering.xml', xml.parse(numberingXml()), TYPES.numbering)
  pkg.add(
    '/word/settings.xml',
    xml.parse(
      `<w:settings xmlns:w="${W}"><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`
    ),
    TYPES.settings
  )
  const main = pkg.rels('/word/document.xml')
  main.add(REL_BASE + 'styles', '/word/styles.xml')
  main.add(REL_BASE + 'numbering', '/word/numbering.xml')
  main.add(REL_BASE + 'settings', '/word/settings.xml')
  addProperties(pkg, title)
}
