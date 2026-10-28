import * as xml from '../xml.ts'
import * as inherit from './inherit.ts'
import { DocsError } from '../errors.ts'
import { nextShapeId, type ShapeEntry } from './shapes.ts'
import { siblings, shapeChildren } from './order.ts'
import { groupFrame, toParent, bounds, writeOffExt } from './geometry.ts'
import type { Element } from '../xml.ts'
import type { Box } from './inherit.ts'
import type { Grouped } from './layout-types.ts'

interface Member {
  shape: ShapeEntry
  box: Box
}

interface GroupSpec {
  id: number
  name: string
  box: Box
}

interface Point {
  x: number
  y: number
}

const { a: A, p: P } = xml.ns
const FULL_TURN = 21600000

function refuse(shape: ShapeEntry, why: string): never {
  throw new DocsError('CANNOT_GROUP', `shape ${shape.id} ${why}`, {
    field: 'shapes',
    got: shape.id
  })
}

function members(shapes: ShapeEntry[]): Member[] {
  if (shapes.length < 2) {
    throw new DocsError('CANNOT_GROUP', 'a group needs at least two shapes', { field: 'shapes' })
  }

  const parent = shapes[0].el.parentNode

  return shapes.map((shape) => {
    if (shape.placeholder !== null) {
      refuse(shape, 'is a placeholder, which cannot be grouped')
    }

    if (shape.el.parentNode !== parent) {
      refuse(shape, 'is not in the same group as the others')
    }

    const box = inherit.explicitBox(shape.el)

    return { shape, box: box ?? refuse(shape, 'has no position of its own') }
  })
}

function groupElement(el: Element, spec: GroupSpec): Element {
  const { id, name, box } = spec
  const doc = xml.owner(el)
  const p = xml.builder(doc, P)
  const a = xml.builder(doc, A)
  const grpSp = p('grpSp')
  const nv = xml.append(grpSp, p('nvGrpSpPr'))
  nv.appendChild(p('cNvPr', { id, name }))
  nv.appendChild(p('cNvGrpSpPr'))
  nv.appendChild(p('nvPr'))
  const xfrm = xml.append(xml.append(grpSp, p('grpSpPr')), a('xfrm'))
  xfrm.appendChild(a('off', { x: box.x, y: box.y }))
  xfrm.appendChild(a('ext', { cx: box.width, cy: box.height }))
  xfrm.appendChild(a('chOff', { x: box.x, y: box.y }))
  xfrm.appendChild(a('chExt', { cx: box.width, cy: box.height }))

  return grpSp
}

export function group(shapes: ShapeEntry[], name: unknown): Grouped {
  const list = members(shapes)
  const order = siblings(list[0].shape.el)
  list.sort((x, y) => order.indexOf(x.shape.el) - order.indexOf(y.shape.el))
  const top = list[list.length - 1].shape.el
  const id = nextShapeId(xml.owner(top))
  const box = bounds(list.map((m) => m.box))
  const label = typeof name === 'string' && name ? name : `Group ${id - 1}`
  const grpSp = groupElement(top, { id, name: label, box })
  top.parentNode?.insertBefore(grpSp, top)
  for (const { shape } of list) {
    grpSp.appendChild(shape.el)
  }

  return { id, ...box, children: list.map((m) => m.shape.id ?? 0) }
}

function turned(box: Box, centre: Point, rot: number): Box {
  const angle = (rot / FULL_TURN) * 2 * Math.PI
  const cx = box.x + box.width / 2 - centre.x
  const cy = box.y + box.height / 2 - centre.y
  const x = centre.x + cx * Math.cos(angle) - cy * Math.sin(angle)
  const y = centre.y + cx * Math.sin(angle) + cy * Math.cos(angle)

  return { ...box, x: Math.round(x - box.width / 2), y: Math.round(y - box.height / 2) }
}

function release(child: Element, frame: ReturnType<typeof groupFrame>, rot: number): void {
  const xfrm = inherit.xfrmOf(child)
  const box = inherit.explicitBox(child)
  if (!xfrm || !box || !frame) {
    return
  }

  const centre = { x: frame.box.x + frame.box.width / 2, y: frame.box.y + frame.box.height / 2 }
  writeOffExt(xfrm, turned(toParent(box, frame), centre, rot))
  if (rot) {
    const own = Number(xfrm.getAttribute('rot') ?? 0)
    xfrm.setAttribute('rot', String((((own + rot) % FULL_TURN) + FULL_TURN) % FULL_TURN))
  }
}

export function ungroup(shape: ShapeEntry): number[] {
  if (shape.el.localName !== 'grpSp') {
    throw new DocsError('NOT_GROUP', `shape ${shape.id} is a ${shape.kind}, not a group`)
  }

  const xfrm = inherit.xfrmOf(shape.el)
  if (xfrm?.getAttribute('flipH') === '1' || xfrm?.getAttribute('flipV') === '1') {
    throw new DocsError(
      'UNSUPPORTED_SHAPE',
      `group ${shape.id} is flipped and cannot be ungrouped yet`
    )
  }

  const frame = groupFrame(shape.el)
  const rot = Number(xfrm?.getAttribute('rot') ?? 0)
  const released = shapeChildren(shape.el)
  for (const child of released) {
    release(child, frame, rot)
    shape.el.parentNode?.insertBefore(child, shape.el)
  }

  xml.remove(shape.el)

  return released.map((el) => Number(xml.descendants(el, P, 'cNvPr')[0]?.getAttribute('id') ?? 0))
}
