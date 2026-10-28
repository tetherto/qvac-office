import * as xml from '../xml.ts'
import { DocsError } from '../errors.ts'
import type { Element } from '../xml.ts'
import type { Reordered, ZOrder } from './layout-types.ts'

const SHAPES = new Set([
  'sp',
  'grpSp',
  'graphicFrame',
  'cxnSp',
  'pic',
  'contentPart',
  'AlternateContent'
])
const ORDERS: readonly ZOrder[] = ['front', 'back', 'forward', 'backward']

export function shapeChildren(parent: Element): Element[] {
  return xml.children(parent).filter((n) => SHAPES.has(n.localName ?? ''))
}

export function siblings(el: Element): Element[] {
  const parent = el.parentNode
  return xml.isElement(parent) ? shapeChildren(parent) : []
}

function target(list: Element[], at: number, order: ZOrder): number {
  switch (order) {
    case 'front':
      return list.length - 1
    case 'back':
      return 0
    case 'forward':
      return Math.min(list.length - 1, at + 1)
    case 'backward':
      return Math.max(0, at - 1)
  }
}

export function reorder(el: Element, order: unknown): Reordered {
  const wanted = ORDERS.find((o) => o === order)
  if (!wanted) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `order must be one of ${ORDERS.join(', ')}, got ${String(order)}`,
      {
        field: 'order',
        expected: ORDERS,
        got: order
      }
    )
  }

  const list = siblings(el)
  const at = list.indexOf(el)
  const to = target(list, at, wanted)
  if (to === at) {
    return { moved: false, position: at }
  }

  const parent = el.parentNode
  const anchor = to > at ? list[to].nextSibling : list[to]
  parent?.insertBefore(el, anchor)

  return { moved: true, position: to }
}
