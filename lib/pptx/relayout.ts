import * as xml from '../xml.ts'
import * as text from '../drawingml/text.ts'
import * as inherit from './inherit.ts'
import { shapesOf, placeholderFrom, nextShapeId, type ShapeEntry } from './shapes.ts'
import { isEmptyPlaceholder } from './overlaps.ts'
import { fitted } from './picture-area.ts'
import { writeOffExt } from './geometry.ts'
import { layoutName } from './read.ts'
import type { Element } from '../xml.ts'
import type { Package } from '../opc/package.ts'
import type { Box } from './inherit.ts'
import type { Relayout } from './layout-types.ts'

export interface Switch {
  pkg: Package
  slide: string
  layout: string
}

interface Pair {
  shape: ShapeEntry
  slot: ShapeEntry
}

const { a: A, p: P } = xml.ns
const REL_LAYOUT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
const CHROME = new Set(['dt', 'ftr', 'sldNum'])
const PH_ATTRIBUTES = ['type', 'orient', 'sz', 'idx']

function role(type: string): string {
  if (type === 'title' || type === 'ctrTitle') return 'title'
  if (type === 'body' || type === 'obj') return 'body'
  return type
}

const TEXT_FALLBACKS: Record<string, string[]> = {
  subTitle: ['subTitle', 'body'],
  body: ['body', 'subTitle']
}
const SOLE_ROLES = new Set(['title', 'body', ...CHROME])

function fallbacks(shape: ShapeEntry): string[] {
  const own = role(shape.placeholder ?? '')
  const text = shape.el.localName === 'sp' ? TEXT_FALLBACKS[own] : undefined
  if (text) {
    return text
  }

  return SOLE_ROLES.has(own) ? [own] : [own, 'obj']
}

function matches(slot: ShapeEntry, wanted: string): boolean {
  const type = slot.placeholder ?? ''
  return wanted === 'obj' ? type === 'obj' : role(type) === wanted
}

function pair(shapes: ShapeEntry[], slots: ShapeEntry[]): Pair[] {
  const free = [...slots]
  return shapes.flatMap((shape) => {
    for (const wanted of fallbacks(shape)) {
      const at = free.findIndex((slot) => matches(slot, wanted))
      if (at !== -1) {
        return [{ shape, slot: free.splice(at, 1)[0] }]
      }
    }

    return []
  })
}

function phOf(el: Element): Element | null {
  const nv = xml.children(el)[0]
  const nvPr = nv ? xml.child(nv, P, 'nvPr') : null
  return nvPr ? xml.child(nvPr, P, 'ph') : null
}

function copyPh(from: Element, to: Element): void {
  const source = phOf(from)
  const target = phOf(to)
  for (const name of PH_ATTRIBUTES) {
    const value = source?.getAttribute(name)
    if (value) {
      target?.setAttribute(name, value)
    } else {
      target?.removeAttribute(name)
    }
  }
}

function slotBox(at: Switch, slot: ShapeEntry): Box | null {
  const probe = placeholderFrom(at.pkg.xml(at.slide), slot, 0)
  return inherit.box({ pkg: at.pkg, slide: at.slide, el: probe })
}

function carry(at: Switch, { shape, slot }: Pair): void {
  copyPh(slot.el, shape.el)
  const xfrm = inherit.xfrmOf(shape.el)
  const current = inherit.explicitBox(shape.el)
  const box = slotBox(at, slot)
  if (!xfrm || !current || !box) {
    return
  }

  if (shape.el.localName === 'sp') {
    xml.remove(xfrm)
  } else {
    writeOffExt(xfrm, shape.kind === 'picture' ? fitted(current, box) : box)
  }
}

function contentText(shape: ShapeEntry): string {
  const txBody = xml.child(shape.el, P, 'txBody')
  return txBody ? text.readText(txBody) : ''
}

function keep(at: Switch, shape: ShapeEntry, box: Box | null): void {
  const props = xml.child(shape.el, P, 'spPr')
  if (!box || !props || inherit.explicitBox(shape.el)) {
    return
  }

  const xfrm = xml.builder(at.pkg.xml(at.slide), A)('xfrm')
  props.insertBefore(xfrm, props.firstChild)
  writeOffExt(xfrm, box)
}

function retarget(at: Switch): void {
  const rels = at.pkg.rels(at.slide)
  for (const rel of rels.byType(REL_LAYOUT)) {
    rels.remove(rel.id)
  }

  rels.add(REL_LAYOUT, at.layout)
}

function addMissing(at: Switch, slots: ShapeEntry[], used: Set<ShapeEntry>): number[] {
  const doc = at.pkg.xml(at.slide)
  const tree = xml.descendants(doc, P, 'spTree')[0]
  return slots
    .filter(
      (slot) => !used.has(slot) && slot.el.localName === 'sp' && !CHROME.has(slot.placeholder ?? '')
    )
    .map((slot) => {
      const id = nextShapeId(doc)
      tree?.appendChild(placeholderFrom(doc, slot, id))
      return id
    })
}

function leftOver(at: Switch, shapes: ShapeEntry[], report: Relayout): void {
  for (const shape of shapes) {
    if (isEmptyPlaceholder(shape)) {
      xml.remove(shape.el)
      report.removed.push(shape.id ?? 0)
      continue
    }

    keep(at, shape, inherit.box({ pkg: at.pkg, slide: at.slide, el: shape.el }))
    report.kept.push({
      id: shape.id ?? 0,
      placeholder: shape.placeholder ?? '',
      text: contentText(shape)
    })
  }
}

export function relayout(at: Switch): Relayout {
  const shapes = shapesOf(at.pkg.xml(at.slide)).filter(
    (s) => s.placeholder !== null && s.group === null
  )
  const slots = shapesOf(at.pkg.xml(at.layout)).filter((s) => s.placeholder !== null)
  const pairs = pair(shapes, slots)
  const paired = new Set(pairs.map((p) => p.shape))
  const report: Relayout = {
    layout: layoutName(at.pkg.xml(at.layout)),
    carried: [],
    kept: [],
    removed: [],
    added: []
  }
  leftOver(
    at,
    shapes.filter((s) => !paired.has(s)),
    report
  )
  retarget(at)
  for (const p of pairs) {
    carry(at, p)
    report.carried.push({
      id: p.shape.id ?? 0,
      from: p.shape.placeholder ?? '',
      to: p.slot.placeholder ?? ''
    })
  }

  report.added = addMissing(at, slots, new Set(pairs.map((p) => p.slot)))

  return report
}
