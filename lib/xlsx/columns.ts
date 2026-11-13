import * as xml from '../xml.ts'
import * as refs from './refs.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Sheet } from './sheet.ts'

export interface ColumnPlan {
  col: number
  count: number
  values: unknown[][] | null
}

export interface ColumnSpan {
  from: string
  to: string
}

const X = xml.ns.x

function validateValues(values: unknown): unknown[][] {
  if (
    !Array.isArray(values) ||
    values.length === 0 ||
    !values.every((column) => Array.isArray(column))
  ) {
    throw new DocsError(
      'INVALID_VALUE',
      'values must be a non-empty array of columns, each an array of values from row 1 down',
      { field: 'values', expected: '[["Header", 1, 2], ...]' }
    )
  }

  const columns = values.filter((column): column is unknown[] => Array.isArray(column))
  if (columns.some((column) => column.length > 100000)) {
    throw new DocsError('INVALID_VALUE', 'a column takes at most 100000 values at a time')
  }

  return columns
}

function countOf(o: Record<string, unknown>, values: unknown[][] | null): number {
  const given = o.count === undefined ? 1 : o.count
  const count = values ? values.length : given
  if (!(typeof count === 'number' && Number.isInteger(count) && count >= 1)) {
    throw new DocsError('INVALID_VALUE', `count must be at least 1, got ${String(o.count)}`)
  }

  if (values && o.count !== undefined && o.count !== count) {
    throw new DocsError(
      'INVALID_VALUE',
      `count is ${String(o.count)} but values holds ${count} columns; give one of them`
    )
  }

  return count
}

export function plan(sheet: Sheet, opts: unknown): ColumnPlan {
  if (!isRecord(opts) || Array.isArray(opts)) {
    throw new DocsError('INVALID_OPTIONS', 'insertColumns takes { column, count?, values? }')
  }

  const col = refs.columnArg(opts.column)
  const values = opts.values === undefined ? null : validateValues(opts.values)
  const count = countOf(opts, values)
  const right = Math.max(-1, ...sheet.cells().map((c) => c.col))
  if (right + count >= refs.MAX_COLUMNS || col + count > refs.MAX_COLUMNS) {
    throw new DocsError(
      'INVALID_RANGE',
      `inserting ${count} column(s) would push cells past column XFD; the last used column is ${refs.column(right)}`
    )
  }

  return { col, count, values }
}

function shiftCols(sheet: Sheet, step: ColumnPlan): void {
  const cols = xml.child(sheet.root, X, 'cols')
  if (!cols) {
    return
  }

  const start = step.col + 1
  for (const el of xml.children(cols, X, 'col')) {
    const min = Number(el.getAttribute('min'))
    const max = Number(el.getAttribute('max'))
    if (min >= start) {
      xml.setAttributes(el, { min: min + step.count, max: max + step.count })
    } else if (max >= start) {
      el.setAttribute('max', String(max + step.count))
    }
  }

  for (const el of xml.children(cols, X, 'col')) {
    if (Number(el.getAttribute('min')) > refs.MAX_COLUMNS) {
      xml.remove(el)
    } else if (Number(el.getAttribute('max')) > refs.MAX_COLUMNS) {
      el.setAttribute('max', String(refs.MAX_COLUMNS))
    }
  }
}

export function shift(sheet: Sheet, step: ColumnPlan): void {
  sheet.number()
  for (const row of sheet.rows()) {
    for (const c of xml.children(row, X, 'c')) {
      const pos = refs.parse(c.getAttribute('r'))
      if (pos.col >= step.col) {
        c.setAttribute('r', refs.cell(pos.row, pos.col + step.count))
      }
    }
  }

  shiftCols(sheet, step)
  sheet.host.state.cellsChanged = true
  sheet.touch()
}

export function fill(sheet: Sheet, step: ColumnPlan): void {
  const styleCol = step.col + step.count
  step.values?.forEach((column, i) => {
    column.forEach((value, row) => {
      if (value === null || value === undefined) return
      const col = step.col + i
      sheet.write(refs.cell(row, col), value)
      const s = sheet.findCell({ row, col: styleCol })?.getAttribute('s')
      if (s) {
        sheet.findCell({ row, col })?.setAttribute('s', s)
      }
    })
  })
}

export function span(step: ColumnPlan): ColumnSpan {
  return { from: refs.column(step.col), to: refs.column(step.col + step.count - 1) }
}
