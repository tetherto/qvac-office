import * as xml from '../xml.ts'
import * as paths from '../opc/paths.ts'
import * as chart from '../drawingml/chart.ts'
import { DocsError } from '../errors.ts'
import { nextShapeId, type ShapeEntry } from './shapes.ts'
import type { Element } from '../xml.ts'
import type { Package } from '../opc/package.ts'
import type { Relationship, Relationships } from '../opc/relationships.ts'

export interface Copy {
  pkg: Package
  from: string
  to: string
}

interface Copier {
  pkg: Package
  seen: Map<string, string>
}

const { a: A, r: R, p: P } = xml.ns
const DEEP = true
const CHART_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'

function freeId(rels: Relationships): string {
  const taken = new Set(rels.all().map((rel) => rel.id))
  let n = taken.size + 1
  while (taken.has(`rId${n}`)) {
    n++
  }

  return `rId${n}`
}

function appendRel(rels: Relationships, attrs: xml.Attributes): void {
  const el = xml.named(rels.doc, 'Relationship', xml.ns.rels)
  rels.root.appendChild(xml.setAttributes(el, attrs))
  rels.dirty = true
}

function numbered(part: string): string {
  const match = /(\d+)(\.\w+)$/.exec(part)
  return match ? part.slice(0, match.index) + '%d' + match[2] : part.replace(/(\.\w+)$/, '%d$1')
}

function copyRel(copier: Copier, name: string, rel: Relationship): void {
  const { pkg } = copier
  const external = rel.external || !rel.part || !pkg.has(rel.part)
  const target =
    external || !rel.part ? rel.target : paths.relative(name, copyPart(copier, rel.part))
  appendRel(pkg.rels(name), {
    Id: rel.id,
    Type: rel.type,
    Target: target,
    TargetMode: rel.external ? 'External' : undefined
  })
}

function copyPart(copier: Copier, part: string): string {
  const { pkg, seen } = copier
  const done = seen.get(part)
  if (done) {
    return done
  }

  const name = pkg.freeName(numbered(part))
  seen.set(part, name)
  pkg.add(name, pkg.bytes(part), pkg.contentType(part) ?? undefined)
  for (const rel of pkg.rels(part).all()) {
    copyRel(copier, name, rel)
  }

  return name
}

function relinkChart(copy: Copy, ref: Element): void {
  const rel = copy.pkg.rels(copy.from).get(ref.getAttributeNS(R, 'id'))
  if (!rel?.part || rel.external || !copy.pkg.has(rel.part)) {
    return
  }

  const part = copyPart({ pkg: copy.pkg, seen: new Map() }, rel.part)
  xml.setNsAttribute(
    ref,
    { namespace: R, local: 'id' },
    copy.pkg.rels(copy.to).add(CHART_REL, part)
  )
}

function relinkOther(copy: Copy, id: string): string {
  const rel = copy.pkg.rels(copy.from).get(id)
  if (!rel) {
    return id
  }

  const rels = copy.pkg.rels(copy.to)
  if (rel.external || !rel.part) {
    const fresh = freeId(rels)
    appendRel(rels, { Id: fresh, Type: rel.type, Target: rel.target, TargetMode: 'External' })
    return fresh
  }

  return rels.add(rel.type, rel.part)
}

function relink(copy: Copy, el: Element): void {
  for (const node of [el, ...xml.descendants(el, '*', '*')]) {
    if (node.namespaceURI === chart.C && node.localName === 'chart') {
      relinkChart(copy, node)
      continue
    }

    const related = Array.from(node.attributes).filter((attr) => attr.namespaceURI === R)
    for (const attr of copy.from === copy.to ? [] : related) {
      node.setAttributeNS(R, attr.name, relinkOther(copy, attr.value))
    }
  }
}

function renumber(copy: Copy, el: Element): void {
  const ids = new Map<string, string>()
  let next = nextShapeId(copy.pkg.xml(copy.to))
  for (const cNvPr of xml.descendants(el, P, 'cNvPr')) {
    ids.set(xml.attr(cNvPr, 'id'), String(next))
    cNvPr.setAttribute('id', String(next++))
  }

  for (const link of [...xml.descendants(el, A, 'stCxn'), ...xml.descendants(el, A, 'endCxn')]) {
    const mapped = ids.get(xml.attr(link, 'id'))
    if (mapped) {
      link.setAttribute('id', mapped)
    } else if (copy.from !== copy.to) {
      xml.remove(link)
    }
  }
}

export function duplicate(copy: Copy, shape: ShapeEntry): Element {
  const isDiagram = xml
    .descendants(shape.el, A, 'graphicData')
    .some((data) => (data.getAttribute('uri') ?? '').endsWith('/diagram'))
  if (isDiagram) {
    throw new DocsError(
      'UNSUPPORTED_SHAPE',
      `shape ${shape.id} holds a SmartArt diagram, which cannot be copied yet`
    )
  }

  const doc = copy.pkg.xml(copy.to)
  const el = doc.importNode(shape.el, DEEP)
  if (!xml.isElement(el)) {
    throw new DocsError('INVALID_XML', 'copying the shape produced no element')
  }

  renumber(copy, el)
  relink(copy, el)

  return el
}
