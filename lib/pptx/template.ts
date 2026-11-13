import * as xml from '../xml.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Package } from '../opc/package.ts'

export interface SlideSize {
  width: number
  height: number
}

export type ThemeColor =
  | 'dark1'
  | 'light1'
  | 'dark2'
  | 'light2'
  | 'accent1'
  | 'accent2'
  | 'accent3'
  | 'accent4'
  | 'accent5'
  | 'accent6'
  | 'hyperlink'
  | 'followedHyperlink'

export interface ThemeFonts {
  heading: string
  body: string
}

export interface CreateOptions {
  size?: '16:9' | '4:3' | SlideSize
  title?: string
  fonts?: Partial<ThemeFonts>
  colors?: Partial<Record<ThemeColor, string>>
}

interface Resolved {
  size: SlideSize
  colors: Record<ThemeColor, string>
  fonts: ThemeFonts
  title: string | undefined
}

interface Rect {
  x: number
  y: number
  cx: number
  cy: number
}

interface Geometry {
  title: Rect
  body: Rect
  date: Rect
  footer: Rect
  number: Rect
  coverTitle: Rect
  coverSubtitle: Rect
  sectionTitle: Rect
  sectionBody: Rect
  left: Rect
  right: Rect
}

interface PlaceholderSpec {
  id: number
  name: string
  ph: Record<string, string | number>
  box?: Rect
  bodyPr?: string
  lstStyle?: string
  prompt: string
}

interface FooterSpec {
  start: number
  boxed: boolean
}

interface LayoutSpec {
  name: string
  type: string
  shapes: string[]
}

interface RootSpec {
  local: string
  attrs: string
  body: string
}

const { a: A, p: P, r: R } = xml.ns
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
const PML = 'application/vnd.openxmlformats-officedocument.presentationml.'
const TYPES = {
  presentation: PML + 'presentation.main+xml',
  master: PML + 'slideMaster+xml',
  layout: PML + 'slideLayout+xml',
  presProps: PML + 'presProps+xml',
  viewProps: PML + 'viewProps+xml',
  tableStyles: PML + 'tableStyles+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
}

export const SIZES: Record<'16:9' | '4:3', SlideSize> = {
  '16:9': { width: 12192000, height: 6858000 },
  '4:3': { width: 9144000, height: 6858000 }
}

const DEFAULT_COLORS: Record<ThemeColor, string> = {
  dark1: '000000',
  light1: 'FFFFFF',
  dark2: '44546A',
  light2: 'E7E6E6',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A5A5A5',
  accent4: 'FFC000',
  accent5: '5B9BD5',
  accent6: '70AD47',
  hyperlink: '0563C1',
  followedHyperlink: '954F72'
}

const DEFAULT_FONTS: ThemeFonts = { heading: 'Calibri Light', body: 'Calibri' }

export const SCHEME_NAMES: Record<ThemeColor, string> = {
  dark1: 'dk1',
  light1: 'lt1',
  dark2: 'dk2',
  light2: 'lt2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hyperlink: 'hlink',
  followedHyperlink: 'folHlink'
}

export const TABLE_STYLE = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'

const FOOTER_BODY_PR =
  '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/>'

function isThemeColor(key: string): key is ThemeColor {
  return key in DEFAULT_COLORS
}

function isFontKey(key: string): key is keyof ThemeFonts {
  return key in DEFAULT_FONTS
}

function isSizeName(value: unknown): value is '16:9' | '4:3' {
  return value === '16:9' || value === '4:3'
}

function sizeSource(value: unknown): Record<string, unknown> | SlideSize | null {
  if (value === undefined || value === '') {
    return SIZES['16:9']
  }

  if (isSizeName(value)) {
    return SIZES[value]
  }

  return isRecord(value) ? value : null
}

function inRange(n: unknown): n is number {
  return typeof n === 'number' && Number.isInteger(n) && n >= 914400 && n <= 51206400
}

export function resolveSize(value: unknown): SlideSize {
  const raw = sizeSource(value)
  const width = raw?.width
  const height = raw?.height
  if (!inRange(width) || !inRange(height)) {
    throw new DocsError(
      'INVALID_OPTIONS',
      'size must be "16:9", "4:3" or { width, height } in EMU between 1 and 56 inches'
    )
  }

  return { width, height }
}

function resolveColors(value: unknown): Record<ThemeColor, string> {
  const colors = { ...DEFAULT_COLORS }
  for (const [key, color] of Object.entries(isRecord(value) ? value : {})) {
    if (!isThemeColor(key)) {
      throw new DocsError(
        'INVALID_OPTIONS',
        `unknown theme color "${key}"; use ${Object.keys(DEFAULT_COLORS).join(', ')}`
      )
    }

    if (typeof color !== 'string' || !/^#?[0-9a-fA-F]{6}$/.test(color)) {
      throw new DocsError(
        'INVALID_OPTIONS',
        `theme color ${key} must be hex RGB, got ${String(color)}`
      )
    }

    colors[key] = color.replace('#', '').toUpperCase()
  }

  return colors
}

function resolveFonts(value: unknown): ThemeFonts {
  const fonts = { ...DEFAULT_FONTS }
  for (const [key, face] of Object.entries(isRecord(value) ? value : {})) {
    if (!isFontKey(key)) {
      throw new DocsError('INVALID_OPTIONS', `unknown theme font "${key}"; use heading or body`)
    }

    if (typeof face !== 'string' || !face.trim()) {
      throw new DocsError('INVALID_OPTIONS', `theme font ${key} must be a typeface name`)
    }

    fonts[key] = face
  }

  return fonts
}

function options(opts: unknown): Resolved {
  const input = isRecord(opts) ? opts : {}
  return {
    size: resolveSize(input.size),
    colors: resolveColors(input.colors),
    fonts: resolveFonts(input.fonts),
    title: typeof input.title === 'string' ? input.title : undefined
  }
}

function escape(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function geometry(size: SlideSize): Geometry {
  const sx = (n: number): number => Math.round((n * size.width) / 12192000)
  const sy = (n: number): number => Math.round((n * size.height) / 6858000)
  const rect = (at: Rect): Rect => ({ x: sx(at.x), y: sy(at.y), cx: sx(at.cx), cy: sy(at.cy) })
  return {
    title: rect({ x: 838200, y: 365125, cx: 10515600, cy: 1325563 }),
    body: rect({ x: 838200, y: 1825625, cx: 10515600, cy: 4351338 }),
    date: rect({ x: 838200, y: 6356350, cx: 2743200, cy: 365125 }),
    footer: rect({ x: 4038600, y: 6356350, cx: 4114800, cy: 365125 }),
    number: rect({ x: 8610600, y: 6356350, cx: 2743200, cy: 365125 }),
    coverTitle: rect({ x: 1524000, y: 1122363, cx: 9144000, cy: 2387600 }),
    coverSubtitle: rect({ x: 1524000, y: 3602038, cx: 9144000, cy: 1655762 }),
    sectionTitle: rect({ x: 831850, y: 1709738, cx: 10515600, cy: 2852737 }),
    sectionBody: rect({ x: 831850, y: 4589463, cx: 10515600, cy: 1500187 }),
    left: rect({ x: 838200, y: 1825625, cx: 5181600, cy: 4351338 }),
    right: rect({ x: 6172200, y: 1825625, cx: 5181600, cy: 4351338 })
  }
}

function xfrm(b: Rect): string {
  return `<a:xfrm><a:off x="${b.x}" y="${b.y}"/><a:ext cx="${b.cx}" cy="${b.cy}"/></a:xfrm>`
}

function run(text: string): string {
  return `<a:p><a:r><a:rPr lang="en-US"/><a:t>${escape(text)}</a:t></a:r></a:p>`
}

function placeholder(spec: PlaceholderSpec): string {
  const attrs = Object.entries(spec.ph)
    .map(([k, v]) => ` ${k}="${v}"`)
    .join('')
  const box = spec.box ? xfrm(spec.box) : ''
  const bodyPr = spec.bodyPr ?? '<a:bodyPr/>'
  const lstStyle = spec.lstStyle ?? '<a:lstStyle/>'

  return `<p:sp><p:nvSpPr><p:cNvPr id="${spec.id}" name="${escape(spec.name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${attrs}/></p:nvPr></p:nvSpPr><p:spPr>${box}</p:spPr><p:txBody>${bodyPr}${lstStyle}${run(spec.prompt)}</p:txBody></p:sp>`
}

function tree(shapes: string[]): string {
  return `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes.join('')}</p:spTree>`
}

function footerStyle(align: string): string {
  return `<a:lstStyle><a:lvl1pPr algn="${align}"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>`
}

function footers(g: Geometry, spec: FooterSpec): string[] {
  const { start, boxed } = spec
  const parts = [
    { name: 'Date Placeholder', ph: { type: 'dt', sz: 'half', idx: 10 }, box: g.date, align: 'l' },
    {
      name: 'Footer Placeholder',
      ph: { type: 'ftr', sz: 'quarter', idx: 11 },
      box: g.footer,
      align: 'ctr'
    },
    {
      name: 'Slide Number Placeholder',
      ph: { type: 'sldNum', sz: 'quarter', idx: 12 },
      box: g.number,
      align: 'r'
    }
  ]

  return parts.map((part, i) =>
    placeholder({
      id: start + i,
      name: part.name,
      ph: part.ph,
      box: boxed ? part.box : undefined,
      bodyPr: boxed ? FOOTER_BODY_PR : undefined,
      lstStyle: boxed ? footerStyle(part.align) : undefined,
      prompt: ''
    })
  )
}

function rootXml(spec: RootSpec): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:${spec.local} xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"${spec.attrs}>${spec.body}</p:${spec.local}>`
}

function titlePlaceholder(prompt = 'Click to edit Master title style'): string {
  return placeholder({ id: 2, name: 'Title 1', ph: { type: 'title' }, prompt })
}

function coverLayout(g: Geometry): LayoutSpec {
  const noBullets =
    '<a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle>'
  return {
    name: 'Title Slide',
    type: 'title',
    shapes: [
      placeholder({
        id: 2,
        name: 'Title 1',
        ph: { type: 'ctrTitle' },
        box: g.coverTitle,
        bodyPr: '<a:bodyPr anchor="b"/>',
        lstStyle:
          '<a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="6000"/></a:lvl1pPr></a:lstStyle>',
        prompt: 'Click to edit Master title style'
      }),
      placeholder({
        id: 3,
        name: 'Subtitle 2',
        ph: { type: 'subTitle', idx: 1 },
        box: g.coverSubtitle,
        lstStyle: noBullets,
        prompt: 'Click to edit Master subtitle style'
      }),
      ...footers(g, { start: 4, boxed: false })
    ]
  }
}

function sectionLayout(g: Geometry): LayoutSpec {
  return {
    name: 'Section Header',
    type: 'secHead',
    shapes: [
      placeholder({
        id: 2,
        name: 'Title 1',
        ph: { type: 'title' },
        box: g.sectionTitle,
        bodyPr: '<a:bodyPr anchor="b"/>',
        lstStyle: '<a:lstStyle><a:lvl1pPr><a:defRPr sz="6000"/></a:lvl1pPr></a:lstStyle>',
        prompt: 'Click to edit Master title style'
      }),
      placeholder({
        id: 3,
        name: 'Text Placeholder 2',
        ph: { type: 'body', idx: 1 },
        box: g.sectionBody,
        lstStyle:
          '<a:lstStyle><a:lvl1pPr marL="0" indent="0"><a:buNone/><a:defRPr sz="2400"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>',
        prompt: 'Click to edit Master text styles'
      }),
      ...footers(g, { start: 4, boxed: false })
    ]
  }
}

function twoContentLayout(g: Geometry): LayoutSpec {
  const content = (id: number, box: Rect): string =>
    placeholder({
      id,
      name: `Content Placeholder ${id - 1}`,
      ph: { sz: 'half', idx: id - 2 },
      box,
      prompt: 'Click to edit Master text styles'
    })
  return {
    name: 'Two Content',
    type: 'twoObj',
    shapes: [
      titlePlaceholder(),
      content(3, g.left),
      content(4, g.right),
      ...footers(g, { start: 5, boxed: false })
    ]
  }
}

function layouts(g: Geometry): LayoutSpec[] {
  const body = placeholder({
    id: 3,
    name: 'Content Placeholder 2',
    ph: { idx: 1 },
    prompt: 'Click to edit Master text styles'
  })
  return [
    coverLayout(g),
    {
      name: 'Title and Content',
      type: 'obj',
      shapes: [titlePlaceholder(), body, ...footers(g, { start: 4, boxed: false })]
    },
    sectionLayout(g),
    twoContentLayout(g),
    {
      name: 'Title Only',
      type: 'titleOnly',
      shapes: [titlePlaceholder(), ...footers(g, { start: 3, boxed: false })]
    },
    { name: 'Blank', type: 'blank', shapes: footers(g, { start: 2, boxed: false }) }
  ]
}

function font(major: string): string {
  return `<a:latin typeface="+${major}-lt"/><a:ea typeface="+${major}-ea"/><a:cs typeface="+${major}-cs"/>`
}

function bodyLevel(sz: number, i: number): string {
  const lvl = i + 1
  return `<a:lvl${lvl}pPr marL="${228600 + 457200 * i}" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="${i === 0 ? 1000 : 500}"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="&#8226;"/><a:defRPr sz="${sz}" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>${font('mn')}</a:defRPr></a:lvl${lvl}pPr>`
}

function otherLevel(lvl: number): string {
  return `<a:lvl${lvl}pPr marL="${457200 * (lvl - 1)}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>${font('mn')}</a:defRPr></a:lvl${lvl}pPr>`
}

function textStyles(): string {
  const body = [2800, 2400, 2000, 1800, 1800, 1800, 1800, 1800, 1800].map(bodyLevel).join('')
  const other = [1, 2, 3, 4, 5, 6, 7, 8, 9].map(otherLevel).join('')
  return `<p:txStyles><p:titleStyle><a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>${font('mj')}</a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle>${body}</p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${other}</p:otherStyle></p:txStyles>`
}

function master(g: Geometry, layoutCount: number): string {
  const shapes = [
    placeholder({
      id: 2,
      name: 'Title Placeholder 1',
      ph: { type: 'title' },
      box: g.title,
      bodyPr:
        '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr>',
      prompt: 'Click to edit Master title style'
    }),
    placeholder({
      id: 3,
      name: 'Text Placeholder 2',
      ph: { type: 'body', idx: 1 },
      box: g.body,
      bodyPr:
        '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr>',
      prompt: 'Click to edit Master text styles'
    }),
    ...footers(g, { start: 4, boxed: true })
  ]
  const ids = Array.from(
    { length: layoutCount },
    (_, i) => `<p:sldLayoutId id="${2147483649 + i}" r:id="rId${i + 1}"/>`
  ).join('')

  return rootXml({
    local: 'sldMaster',
    attrs: '',
    body: `<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>${tree(shapes)}</p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst>${ids}</p:sldLayoutIdLst>${textStyles()}`
  })
}

function theme(colors: Record<ThemeColor, string>, fonts: ThemeFonts): string {
  const scheme = Object.entries(SCHEME_NAMES)
    .map(
      ([key, tag]) =>
        `<a:${tag}><a:srgbClr val="${isThemeColor(key) ? colors[key] : ''}"/></a:${tag}>`
    )
    .join('')
  const face = (name: string): string =>
    `<a:latin typeface="${escape(name)}"/><a:ea typeface=""/><a:cs typeface=""/>`
  const solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
  const line = (w: number): string =>
    `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="${A}" name="QVAC"><a:themeElements><a:clrScheme name="QVAC">${scheme}</a:clrScheme><a:fontScheme name="QVAC"><a:majorFont>${face(fonts.heading)}</a:majorFont><a:minorFont>${face(fonts.body)}</a:minorFont></a:fontScheme><a:fmtScheme name="QVAC"><a:fillStyleLst>${solid}${solid}${solid}</a:fillStyleLst><a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>${solid}${solid}${solid}</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`
}

function addLayouts(pkg: Package, g: Geometry): number {
  const all = layouts(g)
  all.forEach((layout, i) => {
    const name = `/ppt/slideLayouts/slideLayout${i + 1}.xml`
    const body = `<p:cSld name="${escape(layout.name)}">${tree(layout.shapes)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`
    pkg.add(
      name,
      xml.parse(
        rootXml({ local: 'sldLayout', attrs: ` type="${layout.type}" preserve="1"`, body })
      ),
      TYPES.layout
    )
    pkg.rels(name).add(REL_BASE + 'slideMaster', '/ppt/slideMasters/slideMaster1.xml')
    pkg.rels('/ppt/slideMasters/slideMaster1.xml').add(REL_BASE + 'slideLayout', name)
  })

  return all.length
}

function addProps(pkg: Package, title: string | undefined): void {
  pkg.add(
    '/ppt/presProps.xml',
    xml.parse(rootXml({ local: 'presentationPr', attrs: '', body: '' })),
    TYPES.presProps
  )
  pkg.add(
    '/ppt/viewProps.xml',
    xml.parse(
      rootXml({ local: 'viewPr', attrs: '', body: '<p:gridSpacing cx="76200" cy="76200"/>' })
    ),
    TYPES.viewProps
  )
  pkg.add(
    '/ppt/tableStyles.xml',
    xml.parse(`<a:tblStyleLst xmlns:a="${A}" def="${TABLE_STYLE}"/>`),
    TYPES.tableStyles
  )
  const dcTitle = title ? `<dc:title>${escape(title)}</dc:title>` : ''
  pkg.add(
    '/docProps/core.xml',
    xml.parse(
      `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">${dcTitle}</cp:coreProperties>`
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
}

function addRelationships(pkg: Package): void {
  const main = pkg.rels('/ppt/presentation.xml')
  main.add(REL_BASE + 'slideMaster', '/ppt/slideMasters/slideMaster1.xml')
  main.add(REL_BASE + 'theme', '/ppt/theme/theme1.xml')
  main.add(REL_BASE + 'presProps', '/ppt/presProps.xml')
  main.add(REL_BASE + 'viewProps', '/ppt/viewProps.xml')
  main.add(REL_BASE + 'tableStyles', '/ppt/tableStyles.xml')
  pkg.rels('/').add(REL_BASE + 'officeDocument', '/ppt/presentation.xml')
  pkg
    .rels('/')
    .add(
      'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
      '/docProps/core.xml'
    )
  pkg.rels('/').add(REL_BASE + 'extended-properties', '/docProps/app.xml')
}

export function build(pkg: Package, opts: unknown = {}): void {
  const { size, colors, fonts, title } = options(opts)
  const g = geometry(size)
  const presentation = `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldSz cx="${size.width}" cy="${size.height}"/><p:notesSz cx="6858000" cy="9144000"/>`
  pkg.add(
    '/ppt/presentation.xml',
    xml.parse(
      rootXml({ local: 'presentation', attrs: ' saveSubsetFonts="1"', body: presentation })
    ),
    TYPES.presentation
  )
  pkg.add(
    '/ppt/slideMasters/slideMaster1.xml',
    xml.parse(master(g, layouts(g).length)),
    TYPES.master
  )
  addLayouts(pkg, g)
  pkg.add('/ppt/theme/theme1.xml', xml.parse(theme(colors, fonts)), TYPES.theme)
  pkg.rels('/ppt/slideMasters/slideMaster1.xml').add(REL_BASE + 'theme', '/ppt/theme/theme1.xml')
  addProps(pkg, title)
  addRelationships(pkg)
}

export function notesTheme(): Uint8Array {
  return xml.serialize(xml.parse(theme(DEFAULT_COLORS, DEFAULT_FONTS)))
}
