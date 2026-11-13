import * as refs from './refs.ts'
import { DocsError } from '../errors.ts'
import type { Area } from './refs.ts'
import type { Sheet } from './sheet.ts'

export function intersects(a: Area, b: Area): boolean {
  return a.left <= b.right && b.left <= a.right && a.top <= b.bottom && b.top <= a.bottom
}

export function contains(outer: Area, inner: Area): boolean {
  return (
    inner.top >= outer.top &&
    inner.bottom <= outer.bottom &&
    inner.left >= outer.left &&
    inner.right <= outer.right
  )
}

export function size(area: Area): number {
  return (area.bottom - area.top + 1) * (area.right - area.left + 1)
}

export function parse(range: unknown): Area {
  if (typeof range !== 'string') {
    throw new DocsError('INVALID_RANGE', `a range is text like A1:D20, got ${String(range)}`)
  }

  return refs.parseRange(range.replace(/\$/g, ''))
}

export function parseList(sqref: unknown): Area[] {
  if (typeof sqref !== 'string' || !sqref.trim()) {
    throw new DocsError('INVALID_RANGE', 'a range is text like A1:D20 or "A1:A5 C1:C5"')
  }

  return sqref.trim().split(/\s+/).map(parse)
}

export function format(areas: Area[]): string {
  return areas.map((area) => refs.range(area)).join(' ')
}

export function anyIntersect(a: Area[], b: Area[]): boolean {
  return a.some((x) => b.some((y) => intersects(x, y)))
}

export function mergedAreas(sheet: Sheet): Area[] {
  return sheet.merged().map((ref) => refs.parseRange(ref))
}
