import * as xml from '../xml.ts'
import * as refs from './refs.ts'
import * as areas from './areas.ts'
import * as tables from './tables.ts'
import { SHEET_ORDER, type CellEntry, type Sheet } from './sheet.ts'
import { DocsError } from '../errors.ts'
import type { Area } from './refs.ts'
import type { Element } from '../xml.ts'

export interface Merged {
  range: string
  cleared: string[]
}

export interface Unmerged {
  unmerged: string[]
}

const X = xml.ns.x

function mergeList(sheet: Sheet): Element {
  return (
    xml.child(sheet.root, X, 'mergeCells') ??
    xml.insertOrdered(sheet.root, xml.named(sheet.doc, 'mergeCells', X), SHEET_ORDER)
  )
}

function guard(sheet: Sheet, area: Area, range: string): void {
  if (areas.size(area) < 2) {
    throw new DocsError('INVALID_RANGE', `${range} is one cell; a merge covers at least two`)
  }

  const clash = sheet.merged().filter((ref) => areas.intersects(refs.parseRange(ref), area))
  if (clash.length) {
    throw new DocsError(
      'MERGE_OVERLAP',
      `${range} overlaps the merged range ${clash.join(', ')}; unmerge it first`,
      { field: 'range', got: range, candidates: clash }
    )
  }

  const table = tables.on(sheet).find((t) => areas.intersects(t.area, area))
  if (table) {
    throw new DocsError(
      'TABLE_OVERLAP',
      `${range} overlaps table ${table.name} at ${table.ref}; cells in a table cannot be merged`
    )
  }
}

function covered(sheet: Sheet, area: Area): CellEntry[] {
  return sheet.cells().filter(({ el, row, col }) => {
    const inside = areas.contains(area, { top: row, bottom: row, left: col, right: col })
    const corner = row === area.top && col === area.left

    return inside && !corner && (!!el.firstChild || el.hasAttribute('t'))
  })
}

function clearCovered(sheet: Sheet, cells: CellEntry[]): string[] {
  const cleared: string[] = []
  for (const { el, row, col } of cells) {
    for (const n of Array.from(el.childNodes)) {
      el.removeChild(n)
    }

    el.removeAttribute('t')
    cleared.push(refs.cell(row, col))
  }

  if (cleared.length) {
    sheet.host.state.cellsChanged = true
  }

  return cleared
}

export function merge(sheet: Sheet, range: unknown): Merged {
  const area = areas.parse(range)
  const ref = refs.range(area)
  guard(sheet, area, ref)
  const cells = covered(sheet, area)
  for (const { row, col } of cells) {
    sheet.guardWrite({ row, col }, refs.cell(row, col))
  }

  const cleared = clearCovered(sheet, cells)
  const list = mergeList(sheet)
  list.appendChild(xml.setAttributes(xml.named(sheet.doc, 'mergeCell', X), { ref }))
  list.setAttribute('count', String(xml.children(list, X, 'mergeCell').length))
  sheet.touch()

  return { range: ref, cleared }
}

export function unmerge(sheet: Sheet, range: unknown): Unmerged {
  const area = areas.parse(range)
  const list = xml.child(sheet.root, X, 'mergeCells')
  const hits = list
    ? xml
        .children(list, X, 'mergeCell')
        .filter((m) => areas.intersects(refs.parseRange(xml.attr(m, 'ref')), area))
    : []
  if (!list || !hits.length) {
    return { unmerged: [] }
  }

  for (const m of hits) {
    xml.remove(m)
  }

  const count = xml.children(list, X, 'mergeCell').length
  if (count) {
    list.setAttribute('count', String(count))
  } else {
    xml.remove(list)
  }

  sheet.touch()

  return { unmerged: hits.map((m) => xml.attr(m, 'ref')) }
}
