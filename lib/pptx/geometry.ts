import * as xml from '../xml.ts'
import * as inherit from './inherit.ts'
import type { Element } from '../xml.ts'
import type { Box, ShapeInSlide } from './inherit.ts'

export interface GroupFrame {
  box: Box
  child: Box
}

const { a: A, p: P } = xml.ns

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

export function slideBox(shape: ShapeInSlide): Box | null {
  let box = inherit.box(shape)
  for (let up = shape.el.parentNode; box && xml.isElement(up, P, 'grpSp'); up = up.parentNode) {
    const frame = groupFrame(up)
    box = frame ? toParent(box, frame) : box
  }

  return box
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
