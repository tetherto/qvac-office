import * as xml from '../xml.ts'
import * as text from './text.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import { W, REL_BASE, WML, part, styleId, type Host } from './host.ts'
import type { Document, Element } from '../xml.ts'

export type ListKind = 'bullet' | 'number'

export interface ListPlacement {
  numId: string
  level: number
}

const KINDS: readonly string[] = ['bullet', 'number']
const BULLETS = ['•', 'o', '▪']
const BULLET_FONTS = [null, 'Courier New', 'Wingdings']
const NUMBER_FORMATS = ['decimal', 'lowerLetter', 'lowerRoman']
const LEVELS = 9

function numberingRoot(host: Host): Element {
  const name = part(host, {
    rel: REL_BASE + 'numbering',
    type: WML + 'numbering+xml',
    name: '/word/numbering.xml',
    root: `<w:numbering xmlns:w="${W}"/>`
  })

  return host.package.root(name)
}

function levelFormat(abstract: Element, ilvl: number): string | null {
  const lvl = xml.children(abstract, W, 'lvl').find((el) => text.wGet(el, 'ilvl') === String(ilvl))
  return lvl ? text.wGet(xml.child(lvl, W, 'numFmt'), 'val') : null
}

function matches(abstract: Element, kind: ListKind): boolean {
  if (xml.children(abstract, W, 'lvl').length < LEVELS) {
    return false
  }

  const format = levelFormat(abstract, 0)

  return kind === 'bullet' ? format === 'bullet' : format === 'decimal'
}

function findAbstract(root: Element, kind: ListKind): string | null {
  const found = xml.children(root, W, 'abstractNum').find((el) => matches(el, kind))
  return found ? text.wGet(found, 'abstractNumId') : null
}

function levelElement(doc: Document, kind: ListKind, ilvl: number): Element {
  const lvl = text.wEl(doc, 'lvl', { ilvl })
  const bullet = kind === 'bullet'
  lvl.appendChild(text.wEl(doc, 'start', { val: 1 }))
  lvl.appendChild(text.wEl(doc, 'numFmt', { val: bullet ? 'bullet' : NUMBER_FORMATS[ilvl % 3] }))
  lvl.appendChild(text.wEl(doc, 'lvlText', { val: bullet ? BULLETS[ilvl % 3] : `%${ilvl + 1}.` }))
  lvl.appendChild(text.wEl(doc, 'lvlJc', { val: 'left' }))
  xml
    .append(lvl, text.wEl(doc, 'pPr'))
    .appendChild(text.wEl(doc, 'ind', { left: 720 + 360 * ilvl, hanging: 360 }))
  const font = bullet ? BULLET_FONTS[ilvl % 3] : null
  if (font) {
    xml
      .append(lvl, text.wEl(doc, 'rPr'))
      .appendChild(text.wEl(doc, 'rFonts', { ascii: font, hAnsi: font, hint: 'default' }))
  }

  return lvl
}

function maxOf(elements: Element[], attribute: string): number {
  return Math.max(0, ...elements.map((el) => Number(text.wGet(el, attribute)) || 0))
}

function createAbstract(root: Element, kind: ListKind): string {
  const doc = xml.owner(root)
  const abstracts = xml.children(root, W, 'abstractNum')
  const id = String(abstracts.length ? maxOf(abstracts, 'abstractNumId') + 1 : 0)
  const abstract = text.wEl(doc, 'abstractNum', { abstractNumId: id })
  abstract.appendChild(text.wEl(doc, 'multiLevelType', { val: 'hybridMultilevel' }))
  for (let i = 0; i < LEVELS; i++) {
    abstract.appendChild(levelElement(doc, kind, i))
  }

  const after = abstracts[abstracts.length - 1]
  const before = after ? after.nextSibling : firstNum(root)
  root.insertBefore(abstract, before)

  return id
}

function firstNum(root: Element): Element | null {
  return xml.child(root, W, 'num') ?? xml.child(root, W, 'numIdMacAtCleanup')
}

interface NumSpec {
  abstract: string
  kind: ListKind
}

function createNum(root: Element, spec: NumSpec): string {
  const doc = xml.owner(root)
  const id = String(maxOf(xml.children(root, W, 'num'), 'numId') + 1)
  const num = text.wEl(doc, 'num', { numId: id })
  num.appendChild(text.wEl(doc, 'abstractNumId', { val: spec.abstract }))
  for (let i = 0; spec.kind === 'number' && i < LEVELS; i++) {
    xml
      .append(num, text.wEl(doc, 'lvlOverride', { ilvl: i }))
      .appendChild(text.wEl(doc, 'startOverride', { val: 1 }))
  }

  root.insertBefore(num, xml.child(root, W, 'numIdMacAtCleanup'))

  return id
}

function plainNum(root: Element, abstractId: string): string | null {
  const num = xml
    .children(root, W, 'num')
    .find(
      (el) =>
        text.wGet(xml.child(el, W, 'abstractNumId'), 'val') === abstractId &&
        !xml.child(el, W, 'lvlOverride')
    )
  return num ? text.wGet(num, 'numId') : null
}

export class ListAllocator {
  host: Host
  private numbered: string | null
  private bullets: string | null

  constructor(host: Host) {
    this.host = host
    this.numbered = null
    this.bullets = null
  }

  numId(kind: ListKind): string {
    if (kind === 'bullet') {
      this.bullets = this.bullets ?? this.bulletNum()
      return this.bullets
    }

    if (!this.numbered) {
      const root = numberingRoot(this.host)
      this.numbered = this.create(root, { abstract: this.abstract(root, 'number'), kind: 'number' })
    }

    return this.numbered
  }

  end(): void {
    this.numbered = null
  }

  private bulletNum(): string {
    const root = numberingRoot(this.host)
    const abstract = this.abstract(root, 'bullet')
    return plainNum(root, abstract) ?? this.create(root, { abstract, kind: 'bullet' })
  }

  private create(root: Element, spec: NumSpec): string {
    const id = createNum(root, spec)
    this.host.package.touch(this.host.package.related(this.host.main, REL_BASE + 'numbering')[0])
    return id
  }

  private abstract(root: Element, kind: ListKind): string {
    return findAbstract(root, kind) ?? createAbstract(root, kind)
  }
}

export function checkKind(value: unknown): ListKind {
  if (value === 'bullet' || value === 'number') {
    return value
  }

  throw new DocsError('INVALID_TEXT', `list must be "bullet" or "number", got ${String(value)}`, {
    field: 'list',
    expected: KINDS,
    got: value
  })
}

export function checkLevel(value: unknown): number {
  if (value === undefined) {
    return 0
  }

  if (typeof value === 'number' && Number.isInteger(value) && value >= 0 && value < LEVELS) {
    return value
  }

  throw new DocsError(
    'INVALID_TEXT',
    `level must be an integer from 0 to 8, got ${String(value)}`,
    {
      field: 'level',
      expected: '0 to 8',
      got: value
    }
  )
}

export function assign(host: Host, specs: unknown[]): unknown[] {
  const lists = new ListAllocator(host)
  const paragraphStyle = styleId(host, 'List Paragraph')
  return specs.map((spec) => {
    if (!isRecord(spec) || spec.list === undefined) {
      lists.end()
      return spec
    }

    const kind = checkKind(spec.list)
    checkLevel(spec.level)
    const styled = spec.styleId === undefined && paragraphStyle ? { styleId: paragraphStyle } : {}

    return { ...spec, ...styled, numId: lists.numId(kind) }
  })
}

export function place(p: Element, list: ListPlacement): void {
  const doc = xml.owner(p)
  const pPr = text.ensure(p, 'pPr')
  for (const old of xml.children(pPr, W, 'numPr')) {
    xml.remove(old)
  }

  const numPr = xml.insertOrdered(pPr, text.wEl(doc, 'numPr'), text.PARAGRAPH_ORDER)
  numPr.appendChild(text.wEl(doc, 'ilvl', { val: list.level }))
  numPr.appendChild(text.wEl(doc, 'numId', { val: list.numId }))
}

export function apply(p: Element, spec: Record<string, unknown>): void {
  place(p, { numId: String(spec.numId), level: checkLevel(spec.level) })
}

export function isBullet(host: Host, list: ListPlacement): boolean {
  const name = host.package.related(host.main, REL_BASE + 'numbering')[0]
  if (!name) {
    return true
  }

  const root = host.package.root(name)
  const num = xml.children(root, W, 'num').find((el) => text.wGet(el, 'numId') === list.numId)
  const abstractId = num ? text.wGet(xml.child(num, W, 'abstractNumId'), 'val') : ''
  const abstract = xml
    .children(root, W, 'abstractNum')
    .find((el) => text.wGet(el, 'abstractNumId') === abstractId)
  const format = abstract ? levelFormat(abstract, list.level) : null

  return format === null || format === 'bullet' || format === 'none'
}
