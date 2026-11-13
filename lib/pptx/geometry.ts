import * as xml from '../xml.ts'
import * as inherit from './inherit.ts'
import type { Element } from '../xml.ts'
import type { Box, ShapeInSlide } from './inherit.ts'

export interface GroupFrame {
  box: Box
  child: Box
}

export interface Turn {
  rot: number
  flipH: boolean
  flipV: boolean
}

export interface Placed extends Turn {
  box: Box
}

interface Point {
  x: number
  y: number
}

const { a: A, p: P } = xml.ns
const FULL_TURN = 21600000

export function groupFrame(group: Element): GroupFrame | null {
  const box = inherit.explicitBox(group)
  const xfrm = inherit.xfrmOf(group)
  const chOff = xfrm ? xml.child(xfrm, A, 'chOff') : null
  const chExt = xfrm ? xml.child(xfrm, A, 'chExt') : null
  if (!box || !chOff || !chExt) {
    return null
  }

  return {
    box,
    child: {
      x: Number(chOff.getAttribute('x')),
      y: Number(chOff.getAttribute('y')),
      width: Number(chExt.getAttribute('cx')),
      height: Number(chExt.getAttribute('cy'))
    }
  }
}

function ratio(outer: number, inner: number): number {
  return inner === 0 ? 1 : outer / inner
}

export function toParent(box: Box, frame: GroupFrame): Box {
  const sx = ratio(frame.box.width, frame.child.width)
  const sy = ratio(frame.box.height, frame.child.height)

  return {
    x: Math.round(frame.box.x + (box.x - frame.child.x) * sx),
    y: Math.round(frame.box.y + (box.y - frame.child.y) * sy),
    width: Math.round(box.width * sx),
    height: Math.round(box.height * sy)
  }
}

function wholeTurn(rot: number): number {
  return ((rot % FULL_TURN) + FULL_TURN) % FULL_TURN
}

function turned(box: Box, centre: Point, rot: number): Box {
  const angle = (rot / FULL_TURN) * 2 * Math.PI
  const cx = box.x + box.width / 2 - centre.x
  const cy = box.y + box.height / 2 - centre.y
  const x = centre.x + cx * Math.cos(angle) - cy * Math.sin(angle)
  const y = centre.y + cx * Math.sin(angle) + cy * Math.cos(angle)

  return { ...box, x: Math.round(x - box.width / 2), y: Math.round(y - box.height / 2) }
}

function mirrored(box: Box, centre: Point, turn: Turn): Box {
  return {
    ...box,
    x: turn.flipH ? 2 * centre.x - box.x - box.width : box.x,
    y: turn.flipV ? 2 * centre.y - box.y - box.height : box.y
  }
}

export function turnOf(el: Element): Turn {
  const xfrm = inherit.xfrmOf(el)

  return {
    rot: Number(xfrm?.getAttribute('rot') ?? 0),
    flipH: xfrm?.getAttribute('flipH') === '1',
    flipV: xfrm?.getAttribute('flipV') === '1'
  }
}

export function throughGroup(placed: Placed, group: Element): Placed {
  const frame = groupFrame(group)
  if (!frame) {
    return placed
  }

  const turn = turnOf(group)
  const centre = { x: frame.box.x + frame.box.width / 2, y: frame.box.y + frame.box.height / 2 }
  const box = mirrored(toParent(placed.box, frame), centre, turn)
  const own = turn.flipH === turn.flipV ? placed.rot : -placed.rot

  return {
    box: turned(box, centre, turn.rot),
    rot: wholeTurn(turn.rot + own),
    flipH: placed.flipH !== turn.flipH,
    flipV: placed.flipV !== turn.flipV
  }
}

export function slidePlacement(shape: ShapeInSlide): Placed | null {
  const box = inherit.box(shape)
  if (!box) {
    return null
  }

  let placed = { box, ...turnOf(shape.el) }
  for (let up = shape.el.parentNode; xml.isElement(up, P, 'grpSp'); up = up.parentNode) {
    placed = throughGroup(placed, up)
  }

  return placed
}

export function slideBox(shape: ShapeInSlide): Box | null {
  return slidePlacement(shape)?.box ?? null
}

export function intersects(a: Box, b: Box): boolean {
  return a.x < b.x + b.width && b.x < a.x + a.width && a.y < b.y + b.height && b.y < a.y + a.height
}

export function bounds(boxes: Box[]): Box {
  const x = Math.min(...boxes.map((b) => b.x))
  const y = Math.min(...boxes.map((b) => b.y))
  const right = Math.max(...boxes.map((b) => b.x + b.width))
  const bottom = Math.max(...boxes.map((b) => b.y + b.height))

  return { x, y, width: right - x, height: bottom - y }
}

export function writeOffExt(xfrm: Element, box: Box): void {
  const a = xml.builder(xml.owner(xfrm), A)
  let off = xml.child(xfrm, A, 'off')
  if (!off) {
    off = a('off')
    xfrm.insertBefore(off, xfrm.firstChild)
  }

  let ext = xml.child(xfrm, A, 'ext')
  if (!ext) {
    ext = a('ext')
    xfrm.insertBefore(ext, off.nextSibling)
  }

  xml.setAttributes(off, { x: box.x, y: box.y })
  xml.setAttributes(ext, { cx: box.width, cy: box.height })
}

export function writeTurn(xfrm: Element, turn: Turn): void {
  const values = { rot: turn.rot, flipH: turn.flipH ? 1 : 0, flipV: turn.flipV ? 1 : 0 }
  for (const [name, value] of Object.entries(values)) {
    if (value) {
      xfrm.setAttribute(name, String(value))
    } else {
      xfrm.removeAttribute(name)
    }
  }
}
