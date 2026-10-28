import * as xml from '../xml.ts'
import * as text from './text.ts'
import { DocsError } from '../errors.ts'
import type { Package } from '../opc/package.ts'
import type { Styles } from './styles.ts'
import type { Element } from '../xml.ts'

export interface Host {
  package: Package
  main: string
  styles: Styles
}

export interface PartSpec {
  rel: string
  type: string
  name: string
  root: string
}

export interface ListInfo {
  numId: string
  level: number
}

export const W = xml.ns.w
export const R = xml.ns.r
export const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
export const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml'
export const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
export const XMLNS = 'http://www.w3.org/2000/xmlns/'
export const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
export const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'

export const SECTION_ORDER = [
  'headerReference',
  'footerReference',
  'footnotePr',
  'endnotePr',
  'type',
  'pgSz',
  'pgMar',
  'paperSrc',
  'pgBorders',
  'lnNumType',
  'pgNumType',
  'cols',
  'formProt',
  'vAlign',
  'noEndnote',
  'titlePg',
  'textDirection',
  'bidi',
  'rtlGutter',
  'docGrid',
  'printerSettings',
  'sectPrChange'
]

const BLOCKS = new Set(['p', 'tbl', 'sdt'])
const MAX_PARA_ID = 0x7fffffff

export function body(host: Host): Element {
  const found = xml.child(host.package.root(host.main), W, 'body')
  if (!found) {
    throw new DocsError('BROKEN_DOCX', 'the document has no body')
  }

  return found
}

export function blocks(host: Host): Element[] {
  return xml.children(body(host), W).filter((el) => BLOCKS.has(el.localName ?? ''))
}

export function block(host: Host, index: number): Element {
  const all = blocks(host)
  const found = Number.isInteger(index) && index >= 0 ? all[index] : undefined
  if (!found) {
    const range = all.length ? ` (0 to ${all.length - 1})` : ''
    throw new DocsError(
      'BLOCK_NOT_FOUND',
      `block ${index} does not exist; the document has ${all.length} blocks${range}`,
      { field: 'block', got: index, expected: all.length ? `0 to ${all.length - 1}` : 'none' }
    )
  }

  return found
}

export function paragraph(host: Host, index: number): Element {
  const el = block(host, index)
  if (el.localName !== 'p') {
    throw new DocsError('NOT_PARAGRAPH', `block ${index} is not a paragraph`, {
      field: 'block',
      got: index
    })
  }

  return el
}

export function touchMain(host: Host): void {
  host.package.touch(host.main)
}

export function existingPart(host: Host, rel: string): string | null {
  return host.package.related(host.main, rel)[0] ?? null
}

export function part(host: Host, spec: PartSpec): string {
  const existing = existingPart(host, spec.rel)
  if (existing) {
    return existing
  }

  const pkg = host.package
  const name = pkg.has(spec.name) ? pkg.freeName(spec.name.replace(/\.xml$/, '%d.xml')) : spec.name
  pkg.add(name, xml.parse(spec.root), spec.type)
  pkg.rels(host.main).add(spec.rel, name)

  return name
}

export function addExternal(host: Host, type: string, target: string): string {
  const rels = host.package.rels(host.main)
  const taken = new Set(rels.all().map((rel) => rel.id))
  let n = taken.size + 1
  while (taken.has(`rId${n}`)) {
    n++
  }

  const id = `rId${n}`
  const el = xml.named(rels.doc, 'Relationship', xml.ns.rels)
  rels.root.appendChild(
    xml.setAttributes(el, { Id: id, Type: type, Target: target, TargetMode: 'External' })
  )
  rels.dirty = true

  return id
}

export function declare(root: Element, prefix: string, namespace: string): void {
  if (root.lookupNamespaceURI(prefix) !== namespace) {
    root.setAttributeNS(XMLNS, `xmlns:${prefix}`, namespace)
  }
}

export function paraIds(host: Host): Set<string> {
  const out = new Set<string>()
  for (const name of host.package.names()) {
    if (!(host.package.contentType(name) ?? '').startsWith(WML)) {
      continue
    }

    for (const p of xml.descendants(host.package.xml(name), W, 'p')) {
      const id = p.getAttributeNS(W14, 'paraId')
      if (id) {
        out.add(id.toUpperCase())
      }
    }
  }

  return out
}

export function freeParaId(taken: Set<string>): string {
  let n = 0x10000000 + taken.size
  while (taken.has(hex(n)) && n < MAX_PARA_ID) {
    n++
  }

  taken.add(hex(n))

  return hex(n)
}

function hex(n: number): string {
  return n.toString(16).toUpperCase().padStart(8, '0')
}

export function setParaId(p: Element, id: string): void {
  declare(xml.root(xml.owner(p)), 'w14', W14)
  p.setAttributeNS(W14, 'w14:paraId', id)
  p.setAttributeNS(W14, 'w14:textId', '77777777')
}

export function styleIdOf(p: Element): string | null {
  const pPr = xml.child(p, W, 'pPr')
  const pStyle = pPr ? xml.child(pPr, W, 'pStyle') : null
  return pStyle ? text.wGet(pStyle, 'val') || null : null
}

export function headingLevel(host: Host, p: Element): number | null {
  if (p.localName !== 'p') {
    return null
  }

  const pPr = xml.child(p, W, 'pPr')
  const outline = pPr ? xml.child(pPr, W, 'outlineLvl') : null
  if (outline) {
    const level = Number(text.wGet(outline, 'val'))
    return level < 9 ? level + 1 : null
  }

  return host.styles.headingLevel(styleIdOf(p))
}

export function listInfo(host: Host, p: Element): ListInfo | null {
  const pPr = xml.child(p, W, 'pPr')
  const numPr = pPr ? xml.child(pPr, W, 'numPr') : null
  const own = numPr ? text.wGet(xml.child(numPr, W, 'numId'), 'val') : ''
  const numId = own || host.styles.numIdOf(styleIdOf(p)) || ''
  if (!numId || numId === '0') {
    return null
  }

  const ilvl = numPr ? xml.child(numPr, W, 'ilvl') : null

  return { numId, level: ilvl ? Number(text.wGet(ilvl, 'val')) || 0 : 0 }
}

export function styleId(host: Host, name: string, type = 'paragraph'): string | null {
  return host.styles.resolve(name, type)?.id ?? null
}

export function textWidth(host: Host): number {
  const sectPr = xml.child(body(host), W, 'sectPr')
  const pgSz = sectPr ? xml.child(sectPr, W, 'pgSz') : null
  const pgMar = sectPr ? xml.child(sectPr, W, 'pgMar') : null
  const width = pgSz ? Number(text.wGet(pgSz, 'w')) : 12240
  const left = pgMar ? Number(text.wGet(pgMar, 'left')) : 1440
  const right = pgMar ? Number(text.wGet(pgMar, 'right')) : 1440
  const usable = width - left - right
  return usable > 0 ? usable : 9360
}

export function runStyle(host: Host, r: Element, name: string): boolean {
  const id = styleId(host, name, 'character')
  if (!id) {
    return false
  }

  const rPr = text.ensure(r, 'rPr')
  text.wAttr(text.ensure(rPr, 'rStyle', text.RUN_ORDER), 'val', id)

  return true
}

export function nextId(elements: Element[]): number {
  let max = -1
  for (const el of elements) {
    max = Math.max(max, Number(text.wGet(el, 'id')) || 0)
  }

  return max + 1
}

export function requireText(value: unknown, field: string): string {
  if (typeof value !== 'string' || !value.trim()) {
    throw new DocsError('INVALID_TEXT', `${field} must be non-empty text`, {
      field,
      expected: 'string',
      got: value === undefined ? 'undefined' : typeof value
    })
  }

  return value
}
