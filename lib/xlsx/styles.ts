import * as xml from '../xml.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Document, Element } from '../xml.ts'
import type { Package } from '../opc/package.ts'

export type CellAlign = 'left' | 'center' | 'right' | 'justify'

export interface CellStyle {
  format?: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  color?: string
  fill?: string
  size?: number
  font?: string
  align?: CellAlign
  wrap?: boolean
}

const X = xml.ns.x
const TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
const KNOWN = [
  'format',
  'bold',
  'italic',
  'underline',
  'color',
  'fill',
  'size',
  'font',
  'align',
  'wrap'
]
const TOGGLES = ['bold', 'italic', 'underline', 'wrap'] as const
const COLORS = ['color', 'fill'] as const
const FONT_KEYS = ['bold', 'italic', 'underline', 'color', 'size', 'font'] as const
const ALIGN: Record<CellAlign, string> = {
  left: 'left',
  center: 'center',
  right: 'right',
  justify: 'justify'
}
const BUILTIN = new Map<number, string>([
  [0, 'General'],
  [1, '0'],
  [2, '0.00'],
  [3, '#,##0'],
  [4, '#,##0.00'],
  [9, '0%'],
  [10, '0.00%'],
  [11, '0.00E+00'],
  [12, '# ?/?'],
  [13, '# ??/??'],
  [14, 'mm-dd-yy'],
  [15, 'd-mmm-yy'],
  [16, 'd-mmm'],
  [17, 'mmm-yy'],
  [18, 'h:mm AM/PM'],
  [19, 'h:mm:ss AM/PM'],
  [20, 'h:mm'],
  [21, 'h:mm:ss'],
  [22, 'm/d/yy h:mm'],
  [37, '#,##0 ;(#,##0)'],
  [38, '#,##0 ;[Red](#,##0)'],
  [39, '#,##0.00;(#,##0.00)'],
  [40, '#,##0.00;[Red](#,##0.00)'],
  [45, 'mm:ss'],
  [46, '[h]:mm:ss'],
  [47, 'mmss.0'],
  [48, '##0.0E+0'],
  [49, '@']
])
const DATE_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47])
const STYLE_ORDER = [
  'numFmts',
  'fonts',
  'fills',
  'borders',
  'cellStyleXfs',
  'cellXfs',
  'cellStyles',
  'dxfs',
  'tableStyles',
  'colors',
  'extLst'
]
const FONT_ORDER = [
  'b',
  'i',
  'strike',
  'condense',
  'extend',
  'outline',
  'shadow',
  'u',
  'vertAlign',
  'sz',
  'color',
  'name',
  'family',
  'charset',
  'scheme'
]
const DEFAULT = `<styleSheet xmlns="${X}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`

export function isDateCode(code: string): boolean {
  const stripped = code
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/\[(?!h\]|m\]|s\])[^\]]*\]/gi, '')
  if (/^general$/i.test(stripped.trim())) {
    return false
  }

  return /[dmyhs]/i.test(stripped.split(';')[0])
}

function isAlign(value: unknown): value is CellAlign {
  return typeof value === 'string' && value in ALIGN
}

function checkValues(style: Record<string, unknown>): void {
  const { format, size, font, align } = style
  if (format !== undefined && !(typeof format === 'string' && format.length <= 255)) {
    throw new DocsError(
      'INVALID_STYLE',
      'format must be a number format like "0.00" or "yyyy-mm-dd"'
    )
  }

  const toggle = TOGGLES.find((key) => style[key] !== undefined && typeof style[key] !== 'boolean')
  if (toggle) {
    throw new DocsError('INVALID_STYLE', `${toggle} must be true or false`)
  }

  for (const key of COLORS) {
    const value = style[key]
    if (value !== undefined && !(typeof value === 'string' && /^#?[0-9a-fA-F]{6}$/.test(value))) {
      throw new DocsError(
        'INVALID_STYLE',
        `${key} must be a hex RGB value like "1A73E8", got ${String(value)}`
      )
    }
  }

  if (size !== undefined && !(typeof size === 'number' && size >= 1 && size <= 409)) {
    throw new DocsError('INVALID_STYLE', `size is in points, from 1 to 409, got ${String(size)}`)
  }

  if (font !== undefined && !(typeof font === 'string' && font.trim())) {
    throw new DocsError('INVALID_STYLE', 'font must be a typeface name')
  }

  if (align !== undefined && !isAlign(align)) {
    throw new DocsError(
      'INVALID_STYLE',
      `align must be left, center, right or justify, got ${String(align)}`
    )
  }
}

function str(value: unknown): string | undefined {
  return typeof value === 'string' ? value : undefined
}

function bool(value: unknown): boolean | undefined {
  return typeof value === 'boolean' ? value : undefined
}

export function validate(style: unknown): CellStyle {
  if (!isRecord(style) || Array.isArray(style)) {
    throw new DocsError('INVALID_STYLE', 'style must be an object')
  }

  const unknownKey = Object.keys(style).find((key) => !KNOWN.includes(key))
  if (unknownKey) {
    throw new DocsError(
      'INVALID_STYLE',
      `unknown cell style "${unknownKey}"; use ${KNOWN.join(', ')}`
    )
  }

  checkValues(style)

  return {
    format: str(style.format),
    bold: bool(style.bold),
    italic: bool(style.italic),
    underline: bool(style.underline),
    color: str(style.color),
    fill: str(style.fill),
    size: typeof style.size === 'number' ? style.size : undefined,
    font: str(style.font),
    align: isAlign(style.align) ? style.align : undefined,
    wrap: bool(style.wrap)
  }
}

function argb(color: string): string {
  return 'FF' + color.replace('#', '').toUpperCase()
}

function setFontChild(font: Element, local: string, attributes: xml.Attributes): void {
  const doc = xml.owner(font)
  const el =
    xml.child(font, X, local) ?? xml.insertOrdered(font, xml.named(doc, local, X), FONT_ORDER)
  for (const attr of Array.from(el.attributes)) {
    el.removeAttribute(attr.name)
  }

  xml.setAttributes(el, attributes)
}

function dropFontChild(font: Element, local: string): void {
  for (const el of xml.children(font, X, local)) {
    xml.remove(el)
  }
}

function toggleFont(font: Element, local: string, on: boolean | undefined): void {
  if (on === undefined) return
  if (on) {
    setFontChild(font, local, {})
  } else {
    dropFontChild(font, local)
  }
}

function editFont(font: Element, style: CellStyle): void {
  toggleFont(font, 'b', style.bold)
  toggleFont(font, 'i', style.italic)
  toggleFont(font, 'u', style.underline)
  if (style.size !== undefined) {
    setFontChild(font, 'sz', { val: style.size })
  }

  if (style.color !== undefined) {
    setFontChild(font, 'color', { rgb: argb(style.color) })
  }

  if (style.font !== undefined) {
    setFontChild(font, 'name', { val: style.font })
    dropFontChild(font, 'scheme')
  }
}

export class Styles {
  package: Package
  workbookPart: string
  part: string | null

  constructor(pkg: Package, workbookPart: string) {
    this.package = pkg
    this.workbookPart = workbookPart
    this.part = pkg.related(workbookPart, REL)[0] ?? null
  }

  get doc(): Document {
    return this.package.xml(this.ensurePart())
  }

  formatCode(xfIndex: number): string {
    if (!this.part) return 'General'
    const xf = this.xfs()[xfIndex || 0]
    return this.code(xf ? Number(xf.getAttribute('numFmtId') || 0) : 0)
  }

  isDate(xfIndex: number): boolean {
    if (!this.part || !xfIndex) return false
    const xf = this.xfs()[xfIndex]
    if (!xf) {
      return false
    }

    const id = Number(xf.getAttribute('numFmtId') || 0)
    if (DATE_IDS.has(id)) {
      return true
    }

    return id >= 164 && isDateCode(this.code(id))
  }

  derive(xfIndex: number, input: unknown): number {
    const style = validate(input)
    const doc = this.doc
    const xfs = this.xfs()
    const base = xfs[xfIndex || 0] ?? xfs[0]
    const xf = base
      ? xml.clone(base)
      : xml.setAttributes(xml.named(doc, 'xf', X), {
          numFmtId: 0,
          fontId: 0,
          fillId: 0,
          borderId: 0,
          xfId: 0
        })
    this.applyTo(xf, style)
    const key = xf.toString()
    const existing = xfs.findIndex((el) => el.toString() === key)
    if (existing !== -1) {
      return existing
    }

    const list = this.list('cellXfs')
    list.appendChild(xf)
    const count = xml.children(list, X, 'xf').length
    list.setAttribute('count', String(count))
    this.touch()

    return count - 1
  }

  private applyTo(xf: Element, style: CellStyle): void {
    if (style.format !== undefined) {
      xf.setAttribute('numFmtId', String(this.formatId(style.format)))
      xf.setAttribute('applyNumberFormat', '1')
    }

    if (FONT_KEYS.some((k) => style[k] !== undefined)) {
      xf.setAttribute('fontId', String(this.font(Number(xf.getAttribute('fontId') || 0), style)))
      xf.setAttribute('applyFont', '1')
    }

    if (style.fill !== undefined) {
      xf.setAttribute('fillId', String(this.fill(style.fill)))
      xf.setAttribute('applyFill', '1')
    }

    if (style.align !== undefined || style.wrap !== undefined) {
      this.align(xf, style)
    }
  }

  private align(xf: Element, style: CellStyle): void {
    const doc = xml.owner(xf)
    let alignment = xml.child(xf, X, 'alignment')
    if (!alignment) {
      alignment = xml.named(doc, 'alignment', X)
      xf.insertBefore(alignment, xf.firstChild)
    }

    if (style.align !== undefined) {
      alignment.setAttribute('horizontal', ALIGN[style.align])
    }

    if (style.wrap !== undefined) {
      alignment.setAttribute('wrapText', style.wrap ? '1' : '0')
    }

    xf.setAttribute('applyAlignment', '1')
  }

  private ensurePart(): string {
    if (this.part) return this.part
    const part = '/xl/styles.xml'
    this.package.add(part, xml.parse(DEFAULT), TYPE)
    this.package.rels(this.workbookPart).add(REL, part)
    this.part = part
    return part
  }

  private touch(): void {
    this.package.touch(this.ensurePart())
  }

  private xfs(): Element[] {
    const list = xml.child(xml.root(this.doc), X, 'cellXfs')
    return list ? xml.children(list, X, 'xf') : []
  }

  private list(local: string): Element {
    const doc = this.doc
    const top = xml.root(doc)
    const existing = xml.child(top, X, local)
    if (existing) {
      return existing
    }

    return xml.insertOrdered(
      top,
      xml.setAttributes(xml.named(doc, local, X), { count: 0 }),
      STYLE_ORDER
    )
  }

  private code(id: number): string {
    const builtin = BUILTIN.get(id)
    if (builtin !== undefined) {
      return builtin
    }

    const list = xml.child(xml.root(this.doc), X, 'numFmts')
    const el = list
      ? xml.children(list, X, 'numFmt').find((n) => Number(n.getAttribute('numFmtId')) === id)
      : undefined

    return el?.getAttribute('formatCode') || 'General'
  }

  private formatId(code: string): number {
    for (const [id, value] of BUILTIN) {
      if (value === code) return id
    }

    const list = this.list('numFmts')
    const existing = xml.children(list, X, 'numFmt')
    const match = existing.find((n) => n.getAttribute('formatCode') === code)
    if (match) {
      return Number(match.getAttribute('numFmtId'))
    }

    const id = Math.max(163, ...existing.map((n) => Number(n.getAttribute('numFmtId')))) + 1
    list.appendChild(
      xml.setAttributes(xml.named(this.doc, 'numFmt', X), { numFmtId: id, formatCode: code })
    )
    list.setAttribute('count', String(existing.length + 1))
    this.touch()

    return id
  }

  private font(baseId: number, style: CellStyle): number {
    const list = this.list('fonts')
    const fonts = xml.children(list, X, 'font')
    const base = fonts[baseId]
    const font = base ? xml.clone(base) : xml.named(this.doc, 'font', X)
    editFont(font, style)
    return this.dedupe(list, { el: font, siblings: fonts })
  }

  private fill(color: string): number {
    const doc = this.doc
    const list = this.list('fills')
    const fills = xml.children(list, X, 'fill')
    const fill = xml.named(doc, 'fill', X)
    const pattern = xml.append(
      fill,
      xml.setAttributes(xml.named(doc, 'patternFill', X), { patternType: 'solid' })
    )
    pattern.appendChild(xml.setAttributes(xml.named(doc, 'fgColor', X), { rgb: argb(color) }))
    pattern.appendChild(xml.setAttributes(xml.named(doc, 'bgColor', X), { indexed: 64 }))

    return this.dedupe(list, { el: fill, siblings: fills })
  }

  private dedupe(list: Element, entry: Candidate): number {
    const key = entry.el.toString()
    const existing = entry.siblings.findIndex((el) => el.toString() === key)
    if (existing !== -1) {
      return existing
    }

    list.appendChild(entry.el)
    list.setAttribute('count', String(entry.siblings.length + 1))
    this.touch()

    return entry.siblings.length
  }
}

interface Candidate {
  el: Element
  siblings: Element[]
}
