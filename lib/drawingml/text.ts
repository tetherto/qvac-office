import * as xml from '../xml.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Element, Node } from '../xml.ts'
import type { ReadRun } from './runs.ts'

export interface Paragraph {
  text: string
  level?: number
}

export type TextInput = string | (string | Paragraph)[]

export type Alignment = 'left' | 'center' | 'right' | 'justify'

export interface Style {
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  color?: string
  font?: string
  align?: Alignment
}

export interface ReadParagraph {
  text: string
  level: number
  runs?: ReadRun[]
}

export interface Losses {
  hyperlinks?: number
  fields?: number
}

export interface StyleTarget {
  paragraph?: number
}

const A = xml.ns.a
const STYLE_KEYS = ['size', 'bold', 'italic', 'underline', 'color', 'font', 'align']
const TOGGLES = ['bold', 'italic', 'underline'] as const
const ALIGNMENTS: Record<Alignment, string> = {
  left: 'l',
  center: 'ctr',
  right: 'r',
  justify: 'just'
}
const FILLS = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']
const LINKS = ['hlinkClick', 'hlinkMouseOver']
export const RUN_PROPERTY_ORDER = [
  'ln',
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'effectLst',
  'effectDag',
  'highlight',
  'uLnTx',
  'uLn',
  'uFillTx',
  'uFill',
  'latin',
  'ea',
  'cs',
  'sym',
  'hlinkClick',
  'hlinkMouseOver',
  'rtl',
  'extLst'
]

function isAlignment(value: unknown): value is Alignment {
  return typeof value === 'string' && value in ALIGNMENTS
}

function paragraphText(p: Element): string {
  let out = ''
  for (const node of xml.children(p, A)) {
    if (node.localName === 'r' || node.localName === 'fld') {
      out += xml.text(xml.child(node, A, 't'))
    } else if (node.localName === 'br') {
      out += '\n'
    }
  }

  return out
}

function paragraphLevel(p: Element): number {
  const pPr = xml.child(p, A, 'pPr')
  return pPr?.hasAttribute('lvl') ? Number(pPr.getAttribute('lvl')) : 0
}

export function readParagraphs(txBody: Element): ReadParagraph[] {
  return xml
    .children(txBody, A, 'p')
    .map((p) => ({ text: paragraphText(p), level: paragraphLevel(p) }))
}

export function readText(txBody: Element): string {
  return readParagraphs(txBody)
    .map((p) => p.text)
    .join('\n')
}

function normalizeItem(item: unknown): Paragraph {
  if (typeof item === 'string') {
    return { text: item }
  }

  if (
    typeof item !== 'object' ||
    item === null ||
    !('text' in item) ||
    typeof item.text !== 'string'
  ) {
    throw new DocsError('INVALID_TEXT', 'each paragraph must be a string or { text, level }')
  }

  const level = 'level' in item ? item.level : undefined
  if (
    level !== undefined &&
    !(typeof level === 'number' && Number.isInteger(level) && level >= 0 && level <= 8)
  ) {
    throw new DocsError(
      'INVALID_TEXT',
      `paragraph level must be an integer from 0 to 8, got ${String(level)}`
    )
  }

  return { text: item.text, level }
}

export function normalizeParagraphs(input: unknown): Paragraph[] {
  if (typeof input === 'string') {
    return input.split('\n').map((text) => ({ text }))
  }

  if (!Array.isArray(input)) {
    throw new DocsError('INVALID_TEXT', 'text must be a string or an array of paragraphs')
  }

  return input.map(normalizeItem)
}

function unlinked(rPr: Element): Element {
  for (const local of LINKS) {
    for (const el of xml.children(rPr, A, local)) {
      xml.remove(el)
    }
  }

  return rPr
}

export function losses(txBody: Element): Losses {
  const out: Losses = {}
  const hyperlinks = LINKS.reduce((n, local) => n + xml.descendants(txBody, A, local).length, 0)
  const fields = xml.descendants(txBody, A, 'fld').length
  if (hyperlinks) {
    out.hyperlinks = hyperlinks
  }

  if (fields) {
    out.fields = fields
  }

  return out
}

function runPropertiesOf(template: Element | undefined): Element | null {
  if (!template) return null
  const run = xml.child(template, A, 'r')
  if (run) {
    const rPr = xml.child(run, A, 'rPr')
    return rPr ? unlinked(xml.clone(rPr)) : null
  }

  const end = xml.child(template, A, 'endParaRPr')
  if (!end) {
    return null
  }

  const rPr = xml.builder(xml.owner(template), A)('rPr')
  for (const attr of Array.from(end.attributes)) {
    rPr.setAttribute(attr.name, attr.value)
  }

  for (const node of xml.children(end)) {
    rPr.appendChild(xml.clone(node))
  }

  return unlinked(rPr)
}

function paragraphProperties(
  template: Element | undefined,
  paragraph: Paragraph,
  a: xml.Build
): Element | null {
  const existing = template ? xml.child(template, A, 'pPr') : null
  const pPr = existing ? xml.clone(existing) : null
  if (paragraph.level === undefined) {
    return pPr
  }

  const out = pPr ?? a('pPr')
  if (paragraph.level === 0) {
    out.removeAttribute('lvl')
  } else {
    out.setAttribute('lvl', String(paragraph.level))
  }

  return out
}

function appendLine(p: Element, line: string, rPr: Element | null): void {
  const doc = xml.owner(p)
  const a = xml.builder(doc, A)
  const r = xml.append(p, a('r'))
  if (rPr) {
    r.appendChild(xml.clone(rPr))
  }

  xml.append(r, a('t')).appendChild(doc.createTextNode(line))
}

function buildParagraph(
  template: Element | undefined,
  paragraph: Paragraph,
  a: xml.Build
): Element {
  const p = a('p')
  const pPr = paragraphProperties(template, paragraph, a)
  if (pPr) {
    p.appendChild(pPr)
  }

  const rPr = runPropertiesOf(template)
  paragraph.text.split(/\r\n|[\n\v]/).forEach((raw, i) => {
    const line = xml.clean(raw)
    if (i > 0) {
      const br = xml.append(p, a('br'))
      if (rPr) {
        br.appendChild(xml.clone(rPr))
      }
    }

    if (line !== '') {
      appendLine(p, line, rPr)
    }
  })
  const end = template ? xml.child(template, A, 'endParaRPr') : null
  if (end) {
    p.appendChild(unlinked(xml.clone(end)))
  }

  return p
}

export function writeParagraphs(txBody: Element, input: unknown): number {
  const paragraphs = normalizeParagraphs(input)
  if (paragraphs.length === 0) {
    paragraphs.push({ text: '' })
  }

  const a = xml.builder(xml.owner(txBody), A)
  const old = xml.children(txBody, A, 'p')
  const built = paragraphs.map((paragraph, i) =>
    buildParagraph(old[Math.min(i, old.length - 1)], paragraph, a)
  )
  const anchor = old.length ? old[old.length - 1].nextSibling : null
  for (const p of old) {
    xml.remove(p)
  }

  for (const p of built) {
    txBody.insertBefore(p, anchor)
  }

  return built.length
}

export function ensureTextBody(parent: Element, order?: readonly string[]): Element {
  const namespace = parent.namespaceURI ?? A
  const existing = xml.child(parent, namespace, 'txBody')
  if (existing) {
    return existing
  }

  const doc = xml.owner(parent)
  const a = xml.builder(doc, A)
  const txBody = xml.builder(doc, namespace)('txBody')
  for (const local of ['bodyPr', 'lstStyle', 'p']) {
    txBody.appendChild(a(local))
  }

  if (order) {
    xml.insertOrdered(parent, txBody, order)
  } else {
    parent.appendChild(txBody)
  }

  return txBody
}

function isRecordLike(value: unknown): value is Record<string, unknown> {
  return isRecord(value) && !Array.isArray(value)
}

function checkStyleKeys(style: Record<string, unknown>): void {
  const keys = Object.keys(style)
  if (keys.length === 0) {
    throw new DocsError(
      'INVALID_STYLE',
      `style names nothing to change; use ${STYLE_KEYS.join(', ')}`
    )
  }

  const unknownKey = keys.find((key) => !STYLE_KEYS.includes(key))
  if (unknownKey) {
    throw new DocsError(
      'INVALID_STYLE',
      `unknown style "${unknownKey}"; use ${STYLE_KEYS.join(', ')}`
    )
  }
}

function checkStyleValues(style: Record<string, unknown>): void {
  const { size, color, font, align } = style
  if (size !== undefined && !(typeof size === 'number' && size >= 1 && size <= 400)) {
    throw new DocsError('INVALID_STYLE', `size is in points, from 1 to 400, got ${String(size)}`)
  }

  const toggle = TOGGLES.find((key) => style[key] !== undefined && typeof style[key] !== 'boolean')
  if (toggle) {
    throw new DocsError('INVALID_STYLE', `${toggle} must be true or false`)
  }

  if (color !== undefined && !(typeof color === 'string' && /^#?[0-9a-fA-F]{6}$/.test(color))) {
    throw new DocsError(
      'INVALID_STYLE',
      `color must be a hex RGB value like "1A73E8", got ${String(color)}`
    )
  }

  if (align !== undefined && !isAlignment(align)) {
    throw new DocsError(
      'INVALID_STYLE',
      `align must be left, center, right or justify, got ${String(align)}`
    )
  }

  if (font !== undefined && !(typeof font === 'string' && font.trim())) {
    throw new DocsError('INVALID_STYLE', 'font must be a typeface name')
  }
}

function flag(value: unknown): boolean | undefined {
  return typeof value === 'boolean' ? value : undefined
}

export function validateStyle(style: unknown): Style {
  if (!isRecordLike(style)) {
    throw new DocsError('INVALID_STYLE', 'style must be an object')
  }

  checkStyleKeys(style)
  checkStyleValues(style)
  const { size, color, font, align } = style

  return {
    size: typeof size === 'number' ? size : undefined,
    bold: flag(style.bold),
    italic: flag(style.italic),
    underline: flag(style.underline),
    color: typeof color === 'string' ? color : undefined,
    font: typeof font === 'string' ? font : undefined,
    align: isAlignment(align) ? align : undefined
  }
}

function setColor(rPr: Element, color: string): void {
  const a = xml.builder(xml.owner(rPr), A)
  for (const fill of xml.children(rPr, A)) {
    if (FILLS.includes(fill.localName ?? '')) {
      xml.remove(fill)
    }
  }

  const solid = a('solidFill')
  solid.appendChild(a('srgbClr', { val: color.replace('#', '').toUpperCase() }))
  xml.insertOrdered(rPr, solid, RUN_PROPERTY_ORDER)
}

export function applyRunStyle(rPr: Element, style: Style): void {
  if (style.size !== undefined) {
    rPr.setAttribute('sz', String(Math.round(style.size * 100)))
  }

  if (style.bold !== undefined) {
    rPr.setAttribute('b', style.bold ? '1' : '0')
  }

  if (style.italic !== undefined) {
    rPr.setAttribute('i', style.italic ? '1' : '0')
  }

  if (style.underline !== undefined) {
    rPr.setAttribute('u', style.underline ? 'sng' : 'none')
  }

  if (style.color !== undefined) {
    setColor(rPr, style.color)
  }

  if (style.font !== undefined) {
    const latin =
      xml.child(rPr, A, 'latin') ??
      xml.insertOrdered(rPr, xml.builder(xml.owner(rPr), A)('latin'), RUN_PROPERTY_ORDER)
    latin.setAttribute('typeface', style.font)
  }
}

function ensureChild(parent: Element, local: string, before: Node | null): Element {
  const existing = xml.child(parent, A, local)
  if (existing) {
    return existing
  }

  const el = xml.builder(xml.owner(parent), A)(local)
  parent.insertBefore(el, before)

  return el
}

function styleParagraph(p: Element, style: Style, styleRuns: boolean): number {
  if (style.align !== undefined) {
    ensureChild(p, 'pPr', p.firstChild).setAttribute('algn', ALIGNMENTS[style.align])
  }

  let runs = 0
  for (const node of xml.children(p, A)) {
    if (node.localName === 'r' || node.localName === 'fld') {
      if (styleRuns) {
        applyRunStyle(ensureChild(node, 'rPr', node.firstChild), style)
      }

      runs++
    } else if (node.localName === 'br') {
      const rPr = xml.child(node, A, 'rPr')
      if (rPr) {
        applyRunStyle(rPr, style)
      }
    }
  }

  if (styleRuns) {
    applyRunStyle(ensureChild(p, 'endParaRPr', null), style)
  }

  return runs
}

export function applyStyle(txBody: Element, input: unknown, target: StyleTarget = {}): number {
  const style = validateStyle(input)
  const paragraphs = xml.children(txBody, A, 'p')
  let targets = paragraphs
  if (target.paragraph !== undefined) {
    const p = paragraphs[target.paragraph]
    if (!p) {
      throw new DocsError(
        'PARAGRAPH_NOT_FOUND',
        `paragraph ${target.paragraph} does not exist; this text has ${paragraphs.length} (0 to ${paragraphs.length - 1})`
      )
    }

    targets = [p]
  }

  const styleRuns = Object.keys(input ?? {}).some((key) => key !== 'align')

  return targets.reduce((runs, p) => runs + styleParagraph(p, style, styleRuns), 0)
}
