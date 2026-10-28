import * as xml from '../xml.ts'
import * as refs from './refs.ts'
import * as areas from './areas.ts'
import * as formula from './formula.ts'
import * as tables from './tables.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Element } from '../xml.ts'
import type { Area } from './refs.ts'
import type { Ref } from './formula.ts'
import type { Sheet } from './sheet.ts'

export interface SortKey {
  column: string | number
  descending?: boolean
}

export interface SortOptions {
  range: string
  by: SortKey[]
  header?: boolean
}

export interface Sorted {
  range: string
  rows: number
  moved: number
}

interface Key {
  col: number
  descending: boolean
}

interface Plan {
  area: Area
  data: Area
  keys: Key[]
}

interface RowCells {
  row: number
  cells: Map<number, Element>
}

type Sortable = [rank: number, value: number | string | boolean]

const X = xml.ns.x
const BLANK: Sortable = [4, '']

function keysOf(by: unknown, area: Area): Key[] {
  if (!Array.isArray(by) || by.length === 0 || by.length > 64) {
    throw new DocsError('INVALID_OPTIONS', 'by is a list like [{ column: "B", descending: true }]')
  }

  return by.map((key: unknown) => {
    if (!isRecord(key)) {
      throw new DocsError('INVALID_OPTIONS', 'each sort key is { column, descending? }')
    }

    const col = refs.columnArg(key.column)
    if (col < area.left || col > area.right) {
      throw new DocsError(
        'INVALID_RANGE',
        `sort column ${refs.column(col)} is outside ${refs.range(area)}`,
        { field: 'column', got: key.column }
      )
    }

    return { col, descending: key.descending === true }
  })
}

function planOf(opts: unknown): Plan {
  if (!isRecord(opts) || Array.isArray(opts)) {
    throw new DocsError('INVALID_OPTIONS', 'sort takes { range, by, header? }')
  }

  const area = areas.parse(opts.range)
  const header = opts.header === true
  const data = { ...area, top: area.top + (header ? 1 : 0) }
  if (data.top > data.bottom) {
    throw new DocsError('INVALID_RANGE', `${refs.range(area)} has no rows below its header`)
  }

  if (areas.size(data) > 1000000) {
    throw new DocsError('INVALID_RANGE', `sort at most 1000000 cells at a time`)
  }

  return { area, data, keys: keysOf(opts.by, area) }
}

function touchesArea(ref: Ref, area: Area): boolean {
  const end = ref.end ?? ref.start
  const lo = { row: Math.min(ref.start.row ?? 0, end.row ?? 0), col: ref.start.col ?? 0 }
  const rowsAll = ref.start.row === undefined
  const colsAll = ref.start.col === undefined
  const span: Area = {
    top: rowsAll ? 0 : lo.row,
    bottom: rowsAll ? refs.MAX_ROWS - 1 : Math.max(ref.start.row ?? 0, end.row ?? 0),
    left: colsAll ? 0 : Math.min(lo.col, end.col ?? 0),
    right: colsAll ? refs.MAX_COLUMNS - 1 : Math.max(lo.col, end.col ?? 0)
  }

  return areas.intersects(span, area)
}

function selfReferencing(text: string, sheet: Sheet, area: Area): boolean {
  let hit = false
  formula.scan(text, {
    onRef: (ref, { sheet: prefix, original }) => {
      const own = prefix === null || prefix.toLowerCase() === sheet.name.toLowerCase()
      if (own && touchesArea(ref, area)) {
        hit = true
      }

      return original
    }
  })

  return hit
}

function guardFormulas(sheet: Sheet, plan: Plan): void {
  for (const f of xml.descendants(sheet.doc, X, 'f')) {
    const cell = f.parentNode
    if (!xml.isElement(cell)) {
      continue
    }

    const at = sheet.positionOf(cell)
    if (!areas.contains(plan.data, { top: at.row, bottom: at.row, left: at.col, right: at.col })) {
      continue
    }

    const ref = refs.cell(at.row, at.col)
    const kind = f.getAttribute('t')
    if (kind === 'shared' || kind === 'array' || kind === 'dataTable') {
      throw new DocsError(
        'SORT_HAS_FORMULAS',
        `${ref} is part of a ${kind} formula; sorting would break it`
      )
    }

    if (selfReferencing(xml.text(f), sheet, plan.data)) {
      throw new DocsError(
        'SORT_HAS_FORMULAS',
        `${ref} holds =${xml.text(f)}, which refers to cells inside ${refs.range(plan.data)}; sorting would scramble it, so sort values only or move the formula out`,
        { field: 'range', got: refs.range(plan.area) }
      )
    }
  }
}

function guardStructures(sheet: Sheet, plan: Plan): void {
  const ref = refs.range(plan.area)
  const merged = sheet.merged().filter((m) => areas.intersects(refs.parseRange(m), plan.area))
  if (merged.length) {
    throw new DocsError(
      'RANGE_HAS_MERGES',
      `${ref} overlaps the merged range ${merged.join(', ')}`,
      {
        candidates: merged
      }
    )
  }

  for (const table of tables.on(sheet)) {
    const data = {
      ...table.area,
      top: table.area.top + table.header,
      bottom: table.area.bottom - table.totals
    }
    if (areas.intersects(table.area, plan.data) && !areas.contains(data, plan.data)) {
      throw new DocsError(
        'TABLE_OVERLAP',
        `${ref} straddles table ${table.name} at ${table.ref}; sort inside its data rows or outside it`
      )
    }
  }
}

function sortable(sheet: Sheet, el: Element | undefined): Sortable {
  if (!el) return BLANK
  const t = el.getAttribute('t') || 'n'
  const raw = xml.text(xml.child(el, X, 'v'))
  if (t === 's' || t === 'inlineStr' || t === 'str') {
    const value = sheet.value(el)
    return value === null || value === '' ? BLANK : [1, String(value).toLowerCase()]
  }

  if (raw === '') {
    return BLANK
  }

  if (t === 'b') {
    return [2, raw === '1' || raw === 'true']
  }

  if (t === 'e') {
    return [3, raw]
  }

  return [0, Number(raw)]
}

function compare(a: Sortable, b: Sortable, descending: boolean): number {
  if (a[0] !== b[0]) {
    if (a[0] === 4 || b[0] === 4) {
      return a[0] === 4 ? 1 : -1
    }

    return descending ? b[0] - a[0] : a[0] - b[0]
  }

  if (a[1] === b[1]) {
    return 0
  }

  const order = a[1] < b[1] ? -1 : 1

  return descending ? -order : order
}

function collect(sheet: Sheet, data: Area): RowCells[] {
  const rows: RowCells[] = []
  for (let row = data.top; row <= data.bottom; row++) {
    rows.push({ row, cells: new Map() })
  }

  for (const { el, row, col } of sheet.cells()) {
    if (row < data.top || row > data.bottom || col < data.left || col > data.right) {
      continue
    }

    rows[row - data.top].cells.set(col, el)
  }

  return rows
}

function order(sheet: Sheet, plan: Plan, rows: RowCells[]): RowCells[] {
  const keyed = rows.map((r) => ({
    r,
    keys: plan.keys.map((k) => sortable(sheet, r.cells.get(k.col)))
  }))
  keyed.sort((a, b) => {
    for (let i = 0; i < plan.keys.length; i++) {
      const diff = compare(a.keys[i], b.keys[i], plan.keys[i].descending)
      if (diff) {
        return diff
      }
    }

    return a.r.row - b.r.row
  })

  return keyed.map((k) => k.r)
}

function rowFinder(sheet: Sheet): (row: number) => Element {
  const rows = new Map(sheet.numberedRows().map(({ el, r }) => [r, el]))
  return (row) => {
    const found = rows.get(row) ?? sheet.ensureRow(row)
    rows.set(row, found)
    return found
  }
}

function place(target: Element, el: Element, col: number): void {
  const next = xml.children(target, X, 'c').find((c) => refs.parse(c.getAttribute('r')).col > col)
  target.insertBefore(el, next ?? null)
}

function move(sheet: Sheet, plan: Plan, sorted: RowCells[]): number {
  let moved = 0
  const rowAt = rowFinder(sheet)
  for (const { cells } of sorted) {
    for (const el of cells.values()) {
      xml.remove(el)
    }
  }

  sorted.forEach(({ row: from, cells }, i) => {
    const row = plan.data.top + i
    const target = rowAt(row)
    for (const [col, el] of cells) {
      const f = xml.child(el, X, 'f')
      if (row !== from) {
        moved++
        el.setAttribute('r', refs.cell(row, col))
      }

      if (f && row !== from) {
        xml.setText(f, formula.shift(xml.text(f), { rows: row - from, cols: 0 }))
        sheet.host.state.formulas = true
      }

      place(target, el, col)
    }
  })

  return moved
}

export function sort(sheet: Sheet, opts: unknown): Sorted {
  const plan = planOf(opts)
  guardStructures(sheet, plan)
  guardFormulas(sheet, plan)
  sheet.number()
  const rows = collect(sheet, plan.data)
  const moved = move(sheet, plan, order(sheet, plan, rows))
  sheet.host.state.cellsChanged ||= moved > 0
  sheet.touch()

  return { range: refs.range(plan.area), rows: rows.length, moved }
}
