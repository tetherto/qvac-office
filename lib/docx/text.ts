import * as xml from '../xml.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Document, Element } from '../xml.ts'

export type Alignment = 'left' | 'center' | 'right' | 'justify'

export interface RunFormat {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  size?: number
  color?: string
  font?: string
}

export interface Format extends RunFormat {
  align?: Alignment
}

export interface Run extends RunFormat {
  text: string
}

export interface ParagraphSpec {
  text?: string
  runs?: (string | Run)[]
  style?: string
  styleId?: string
  align?: Alignment
  pageBreak?: boolean
  level?: number
  list?: 'bullet' | 'number'
}

export interface Removed {
  hyperlinks?: number
  images?: number
  fields?: number
  contentControls?: number
  equations?: number
  objects?: number
  breaks?: number
}

export type WAttributes = Record<string, string | number | undefined | null>

export const W = xml.ns.w
const XML_NS = 'http://www.w3.org/XML/1998/namespace'
const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

export const RUN_ORDER = [
  'rStyle',
  'rFonts',
  'b',
  'bCs',
  'i',
  'iCs',
  'caps',
  'smallCaps',
  'strike',
  'dstrike',
  'outline',
  'shadow',
  'emboss',
  'imprint',
  'noProof',
  'snapToGrid',
  'vanish',
  'webHidden',
  'color',
  'spacing',
  'w',
  'kern',
  'position',
  'sz',
  'szCs',
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
  'oMath'
]

export const PARAGRAPH_ORDER = [
  'pStyle',
  'keepNext',
  'keepLines',
  'pageBreakBefore',
  'framePr',
  'widowControl',
  'numPr',
  'suppressLineNumbers',
  'pBdr',
  'shd',
  'tabs',
  'suppressAutoHyphens',
  'kinsoku',
  'wordWrap',
  'overflowPunct',
  'topLinePunct',
  'autoSpaceDE',
  'autoSpaceDN',
  'bidi',
  'adjustRightInd',
  'snapToGrid',
  'spacing',
  'ind',
  'contextualSpacing',
  'mirrorIndents',
  'suppressOverlap',
  'jc',
  'textDirection',
  'textAlignment',
  'textboxTightWrap',
  'outlineLvl',
  'divId',
  'cnfStyle',
  'rPr',
  'sectPr',
  'pPrChange'
]

const CONTAINERS = new Set([
  'hyperlink',
  'ins',
  'moveTo',
  'smartTag',
  'customXml',
  'fldSimple',
  'dir',
  'bdo',
  'sdtContent'
])
const KEPT = new Set(['bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd'])
const REFERENCES = ['commentReference', 'footnoteReference', 'endnoteReference']
export const ALIGNMENTS: Record<Alignment, string> = {
  left: 'left',
  center: 'center',
  right: 'right',
  justify: 'both'
}
export const FORMAT_KEYS: readonly string[] = [
  'bold',
  'italic',
  'underline',
  'size',
  'color',
  'font',
  'align'
]
export const RUN_KEYS: readonly string[] = FORMAT_KEYS.filter((k) => k !== 'align')
const TOGGLES = ['bold', 'italic', 'underline'] as const
const LINK_FORMAT = ['rStyle', 'color', 'u']
const OUTSIDE_TEXT = new Set(['hyperlink', 'fldSimple', 'sdt'])
const REVISIONS = ['rPrChange', 'pPrChange', 'ins', 'del', 'moveFrom', 'moveTo']
const TRACKED = ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange']
const BREAKS = new Set(['\n', '\v'])

export function wGet(el: Element | null, name: string): string {
  return el ? el.getAttributeNS(W, name) || '' : ''
}

function runText(r: Element): string {
  let out = ''
  for (const n of xml.children(r, W)) {
    const name = n.localName
    if (name === 't') {
      out += xml.text(n)
    } else if (name === 'tab') {
      out += '\t'
    } else if (name === 'br') {
      out += wGet(n, 'type') === 'page' ? '' : '\n'
    } else if (name === 'cr') {
      out += '\n'
    } else if (name === 'noBreakHyphen') {
      out += '-'
    }
  }

  return out
}

export function paragraphText(el: Element): string {
  let out = ''
  for (const n of xml.children(el, W)) {
    const name = n.localName ?? ''
    if (name === 'r') {
      out += runText(n)
    } else if (name === 'sdt') {
      const content = xml.child(n, W, 'sdtContent')
      if (content) {
        out += paragraphText(content)
      }
    } else if (CONTAINERS.has(name)) {
      out += paragraphText(n)
    }
  }

  return out
}

export function blockText(el: Element): string {
  if (el.localName === 'p') {
    return paragraphText(el)
  }

  return xml.descendants(el, W, 'p').map(paragraphText).join('\n')
}

export function hasPageBreak(p: Element): boolean {
  return xml.descendants(p, W, 'br').some((br) => wGet(br, 'type') === 'page')
}

export function wAttr(el: Element, name: string, value: string | number): void {
  el.setAttributeNS(W, `${el.prefix || 'w'}:${name}`, String(value))
}

export function wEl(doc: Document, local: string, attributes: WAttributes = {}): Element {
  const el = xml.builder(doc, W)(local)
  for (const [key, value] of Object.entries(attributes)) {
    if (value !== undefined && value !== null) {
      wAttr(el, key, value)
    }
  }

  return el
}

function isAlignment(value: unknown): value is Alignment {
  return typeof value === 'string' && value in ALIGNMENTS
}

function checkKeys(format: Record<string, unknown>, known: readonly string[]): void {
  const keys = Object.keys(format)
  if (keys.length === 0) {
    throw new DocsError('INVALID_STYLE', `format names nothing to change; use ${known.join(', ')}`)
  }

  const unknownKey = keys.find((key) => !known.includes(key))
  if (unknownKey) {
    throw new DocsError('INVALID_STYLE', `unknown format "${unknownKey}"; use ${known.join(', ')}`)
  }
}

function checkValues(format: Record<string, unknown>): void {
  const { size, color, font, align } = format
  if (size !== undefined && !(typeof size === 'number' && size >= 1 && size <= 1638)) {
    throw new DocsError('INVALID_STYLE', `size is in points, from 1 to 1638, got ${String(size)}`)
  }

  const toggle = TOGGLES.find(
    (key) => format[key] !== undefined && typeof format[key] !== 'boolean'
  )
  if (toggle) {
    throw new DocsError('INVALID_STYLE', `${toggle} must be true or false`)
  }

  if (color !== undefined && !(typeof color === 'string' && /^#?[0-9a-fA-F]{6}$/.test(color))) {
    throw new DocsError(
      'INVALID_STYLE',
      `color must be a hex RGB value like "1A73E8", got ${String(color)}`
    )
  }

  if (font !== undefined && !(typeof font === 'string' && font.trim())) {
    throw new DocsError('INVALID_STYLE', 'font must be a typeface name')
  }

  if (align !== undefined && !isAlignment(align)) {
    throw new DocsError(
      'INVALID_STYLE',
      `align must be left, center, right or justify, got ${String(align)}`
    )
  }
}

function flag(value: unknown): boolean | undefined {
  return typeof value === 'boolean' ? value : undefined
}

export function validateFormat(format: unknown, known: readonly string[] = FORMAT_KEYS): Format {
  if (!isRecord(format) || Array.isArray(format)) {
    throw new DocsError('INVALID_STYLE', 'format must be an object')
  }

  checkKeys(format, known)
  checkValues(format)
  const { size, color, font, align } = format
  const out: Format = {}
  if (typeof size === 'number') {
    out.size = size
  }

  if (flag(format.bold) !== undefined) {
    out.bold = flag(format.bold)
  }

  if (flag(format.italic) !== undefined) {
    out.italic = flag(format.italic)
  }

  if (flag(format.underline) !== undefined) {
    out.underline = flag(format.underline)
  }

  if (typeof color === 'string') {
    out.color = color
  }

  if (typeof font === 'string') {
    out.font = font
  }

  if (isAlignment(align)) {
    out.align = align
  }

  return out
}

function ensureRunChild(rPr: Element, local: string): Element {
  return xml.child(rPr, W, local) ?? xml.insertOrdered(rPr, wEl(xml.owner(rPr), local), RUN_ORDER)
}

function setToggle(rPr: Element, local: string, value: boolean): void {
  const el = ensureRunChild(rPr, local)
  if (value) {
    el.removeAttributeNS(W, 'val')
  } else {
    wAttr(el, 'val', '0')
  }
}

function setValue(
  rPr: Element,
  local: string,
  attributes: Record<string, string | number>
): Element {
  const el = ensureRunChild(rPr, local)
  for (const [key, value] of Object.entries(attributes)) {
    wAttr(el, key, value)
  }

  return el
}

function applyColor(rPr: Element, color: string): void {
  const el = setValue(rPr, 'color', { val: color.replace('#', '').toUpperCase() })
  for (const attr of ['themeColor', 'themeTint', 'themeShade']) {
    el.removeAttributeNS(W, attr)
  }
}

function applyFont(rPr: Element, font: string): void {
  const el = setValue(rPr, 'rFonts', { ascii: font, hAnsi: font, cs: font, eastAsia: font })
  for (const attr of ['asciiTheme', 'hAnsiTheme', 'cstheme', 'eastAsiaTheme']) {
    el.removeAttributeNS(W, attr)
  }
}

export function applyRunFormat(rPr: Element, format: RunFormat): void {
  if (format.bold !== undefined) {
    setToggle(rPr, 'b', format.bold)
    setToggle(rPr, 'bCs', format.bold)
  }

  if (format.italic !== undefined) {
    setToggle(rPr, 'i', format.italic)
    setToggle(rPr, 'iCs', format.italic)
  }

  if (format.underline !== undefined) {
    setValue(rPr, 'u', { val: format.underline ? 'single' : 'none' })
  }

  if (format.size !== undefined) {
    const half = Math.round(format.size * 2)
    setValue(rPr, 'sz', { val: half })
    setValue(rPr, 'szCs', { val: half })
  }

  if (format.color !== undefined) {
    applyColor(rPr, format.color)
  }

  if (format.font !== undefined) {
    applyFont(rPr, format.font)
  }
}

export function ensure(parent: Element, local: string, order?: readonly string[]): Element {
  const existing = xml.child(parent, W, local)
  if (existing) {
    return existing
  }

  const el = wEl(xml.owner(parent), local)
  if (order) {
    xml.insertOrdered(parent, el, order)
  } else {
    parent.insertBefore(el, parent.firstChild)
  }

  return el
}

export function setAlign(p: Element, align: Alignment): void {
  const pPr = ensure(p, 'pPr')
  wAttr(ensure(pPr, 'jc', PARAGRAPH_ORDER), 'val', ALIGNMENTS[align])
}

function normalizeRun(run: unknown, i: number): Run {
  if (typeof run === 'string') {
    return { text: run }
  }

  if (!isRecord(run) || typeof run.text !== 'string') {
    throw new DocsError('INVALID_TEXT', `run ${i} must be a string or { text, bold, italic, ... }`)
  }

  const { text, ...rest } = run
  const format = Object.keys(rest).length ? validateFormat(rest, RUN_KEYS) : {}

  return { ...format, text }
}

export function normalizeRuns(input: unknown): Run[] {
  if (typeof input === 'string') {
    return [{ text: input }]
  }

  if (isRecord(input) && Array.isArray(input.runs)) {
    return input.runs.map(normalizeRun)
  }

  if (isRecord(input) && typeof input.text === 'string') {
    return [{ text: input.text }]
  }

  throw new DocsError('INVALID_TEXT', 'text must be a string, { text } or { runs: [...] }')
}

function withoutRevisions(el: Element): Element {
  for (const local of REVISIONS) {
    for (const node of xml.descendants(el, W, local)) {
      xml.remove(node)
    }
  }

  return el
}

function insideOf(r: Element, p: Element): string | null {
  for (let el = r.parentNode; el && el !== p; el = el.parentNode) {
    if (xml.isElement(el, W) && OUTSIDE_TEXT.has(el.localName ?? '')) {
      return el.localName
    }
  }

  return null
}

function withoutLinkFormat(rPr: Element): Element {
  for (const local of LINK_FORMAT) {
    for (const el of xml.children(rPr, W, local)) {
      xml.remove(el)
    }
  }

  return rPr
}

function templateRunProperties(p: Element | null): Element | null {
  if (!p) return null
  const runs = xml.descendants(p, W, 'r').filter((r) => xml.child(r, W, 't'))
  const plain = runs.find((r) => insideOf(r, p) === null)
  const run = plain ?? runs[0]
  if (run) {
    const rPr = xml.child(run, W, 'rPr')
    if (!rPr) {
      return null
    }

    const copy = withoutRevisions(xml.clone(rPr))

    return plain || insideOf(run, p) !== 'hyperlink' ? copy : withoutLinkFormat(copy)
  }

  const pPr = xml.child(p, W, 'pPr')
  const mark = pPr ? xml.child(pPr, W, 'rPr') : null

  return mark ? withoutRevisions(xml.clone(mark)) : null
}

function runFormatOf(run: Run): RunFormat {
  return Object.fromEntries(Object.entries(run).filter(([key]) => key !== 'text'))
}

function appendPiece(r: Element, piece: string): void {
  const doc = xml.owner(r)
  if (BREAKS.has(piece)) {
    r.appendChild(wEl(doc, 'br'))
    return
  }

  if (piece === '\t') {
    r.appendChild(wEl(doc, 'tab'))
    return
  }

  const t = wEl(doc, 't')
  if (/^\s|\s$/.test(piece)) {
    t.setAttributeNS(XML_NS, 'xml:space', 'preserve')
  }

  t.appendChild(doc.createTextNode(piece))
  r.appendChild(t)
}

function buildRun(doc: Document, run: Run, base: Element | null): Element {
  const format = runFormatOf(run)
  const r = wEl(doc, 'r')
  let rPr = base ? xml.clone(base) : null
  if (Object.keys(format).length) {
    rPr = rPr ?? wEl(doc, 'rPr')
    applyRunFormat(rPr, format)
  }

  if (rPr?.firstChild) {
    r.appendChild(rPr)
  }

  for (const piece of run.text.replace(/\r\n?/g, '\n').split(/(\n|\v|\t)/)) {
    const value = BREAKS.has(piece) ? piece : xml.clean(piece)
    if (value !== '') {
      appendPiece(r, value)
    }
  }

  return r
}

export function assertUntracked(p: Element, where: string): void {
  if (TRACKED.some((local) => xml.descendants(p, W, local).length)) {
    throw new DocsError(
      'TRACKED_CHANGES',
      `${where} has tracked changes; rewriting it would accept them silently, so accept or reject them in Word first`
    )
  }
}

function inFallback(el: Element): boolean {
  for (let node = el.parentNode; node; node = node.parentNode) {
    if (xml.isElement(node, MC, 'Fallback')) {
      return true
    }
  }

  return false
}

function count(p: Element, namespace: string, local: string): number {
  return xml.descendants(p, namespace, local).filter((el) => !inFallback(el)).length
}

function equations(p: Element): number {
  const inline = xml
    .descendants(p, M, 'oMath')
    .filter((el) => !xml.isElement(el.parentNode, M, 'oMathPara'))
  return count(p, M, 'oMathPara') + inline.filter((el) => !inFallback(el)).length
}

function losses(p: Element): Removed {
  const begins = xml.descendants(p, W, 'fldChar').filter((f) => wGet(f, 'fldCharType') === 'begin')
  const breaks = xml.descendants(p, W, 'br').filter((br) => {
    const type = wGet(br, 'type')
    return type === 'page' || type === 'column'
  })
  const found: Removed = {
    hyperlinks: count(p, W, 'hyperlink'),
    images: count(p, W, 'drawing') + count(p, W, 'pict'),
    fields: count(p, W, 'fldSimple') + begins.length,
    contentControls: count(p, W, 'sdt'),
    equations: equations(p),
    objects: count(p, W, 'object'),
    breaks: breaks.length
  }

  return Object.fromEntries(Object.entries(found).filter(([, n]) => n > 0))
}

export function replaceContent(p: Element, input: unknown): Removed {
  const doc = xml.owner(p)
  const runs = normalizeRuns(input)
  const base = templateRunProperties(p)
  const removed = losses(p)
  const kept: Element[] = []
  const references: Element[] = []
  for (const n of xml.children(p)) {
    const name = n.localName ?? ''
    if (name === 'pPr') {
      continue
    }

    if (KEPT.has(name)) {
      kept.push(n)
    } else if (name === 'r' && REFERENCES.some((ref) => xml.child(n, W, ref))) {
      references.push(n)
    }

    xml.remove(n)
  }

  const starts = kept.filter((n) => (n.localName ?? '').endsWith('Start'))
  const ends = kept.filter((n) => !(n.localName ?? '').endsWith('Start'))
  const built = runs.map((run) => buildRun(doc, run, base))
  for (const n of [...starts, ...built, ...ends, ...references]) {
    p.appendChild(n)
  }

  return removed
}

function specOf(input: unknown): Record<string, unknown> {
  if (typeof input === 'string') {
    return { text: input }
  }

  if (!isRecord(input) || Array.isArray(input)) {
    throw new DocsError('INVALID_TEXT', 'each paragraph must be a string or an object')
  }

  return input
}

function copyTemplate(p: Element, template: Element): Element | null {
  const pPr = xml.child(template, W, 'pPr')
  if (pPr) {
    const copy = withoutRevisions(xml.clone(pPr))
    for (const el of xml.children(copy, W, 'sectPr')) {
      xml.remove(el)
    }

    p.appendChild(copy)
  }

  return templateRunProperties(template)
}

export function buildParagraph(doc: Document, input: unknown, template: Element | null): Element {
  const p = wEl(doc, 'p')
  const spec = specOf(input)
  if (spec.pageBreak) {
    xml.append(p, wEl(doc, 'r')).appendChild(wEl(doc, 'br', { type: 'page' }))
    return p
  }

  const base = template && spec.style === undefined ? copyTemplate(p, template) : null
  if (typeof spec.styleId === 'string' && spec.styleId) {
    wAttr(ensure(ensure(p, 'pPr'), 'pStyle', PARAGRAPH_ORDER), 'val', spec.styleId)
  }

  if (spec.align) {
    const { align } = validateFormat({ align: spec.align })
    if (align) {
      setAlign(p, align)
    }
  }

  for (const run of normalizeRuns(spec)) {
    p.appendChild(buildRun(doc, run, base))
  }

  return p
}
