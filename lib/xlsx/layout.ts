import * as xml from '../xml.ts'
import * as refs from './refs.ts'
import * as units from '../units.ts'
import { SHEET_ORDER, type Sheet } from './sheet.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Element } from '../xml.ts'

export type RowSpanInput = number | [number, number]
export type ColumnSpanInput = string | number | [string | number, string | number]

export interface RowHeightOptions {
  row: number
  height: number | string
}

export interface RowHeight {
  row: number
  height: number
}

export interface HideOptions {
  rows?: RowSpanInput
  columns?: ColumnSpanInput
  hidden?: boolean
}

export interface Hidden {
  rows: number
  columns: number
  hidden: boolean
}

export interface FreezeOptions {
  rows?: number
  columns?: number
}

export interface Frozen {
  rows: number
  columns: number
  topLeftCell: string | null
}

const X = xml.ns.x
const EMU_PER_POINT = 12700
const MAX_HEIGHT = 409
const DEFAULT_WIDTH = 9.140625
const VIEW_ORDER = ['pane', 'selection', 'pivotSelection', 'extLst']

function record(opts: unknown, method: string, shape: string): Record<string, unknown> {
  if (!isRecord(opts) || Array.isArray(opts)) {
    throw new DocsError('INVALID_OPTIONS', `${method} takes ${shape}`)
  }

  return opts
}

function pointsOf(value: unknown): number {
  const raw = units.length(value, 'height')
  const height = typeof value === 'string' && typeof raw === 'number' ? raw / EMU_PER_POINT : raw
  if (!(typeof height === 'number' && height >= 0 && height <= MAX_HEIGHT)) {
    throw new DocsError(
      'INVALID_VALUE',
      `height is in points, from 0 to ${MAX_HEIGHT}, or a length like "0.5in", got ${String(value)}`,
      { field: 'height', got: value, expected: `0 to ${MAX_HEIGHT} points` }
    )
  }

  return Math.round(height * 100) / 100
}

export function setRowHeight(sheet: Sheet, opts: unknown): RowHeight {
  const o = record(opts, 'setRowHeight', '{ row, height }')
  const r = refs.rowArg(o.row)
  const height = pointsOf(o.height)
  xml.setAttributes(sheet.ensureRow(r), { ht: height, customHeight: 1 })
  sheet.touch()

  return { row: r + 1, height }
}

function rowSpan(value: unknown): [number, number] {
  const [a, b] = Array.isArray(value) && value.length === 2 ? value : [value, value]
  const from = refs.rowArg(a, 'rows')
  const to = refs.rowArg(b, 'rows')
  if (to - from >= 100000 || to < from) {
    throw new DocsError(
      'INVALID_RANGE',
      `rows is a row number or [from, to] with from <= to, at most 100000 rows, got ${JSON.stringify(value)}`
    )
  }

  return [from, to]
}

function columnSpan(value: unknown): [number, number] {
  const [a, b] = Array.isArray(value) && value.length === 2 ? value : [value, value]
  const from = refs.columnArg(a)
  const to = refs.columnArg(b)
  if (to < from) {
    throw new DocsError(
      'INVALID_RANGE',
      `columns is a column or [from, to] with from <= to, got ${JSON.stringify(value)}`
    )
  }

  return [from, to]
}

function rowRange(sheet: Sheet, span: [number, number], create: boolean): Element[] {
  sheet.number()
  const data = sheet.sheetData()
  const rows = sheet.numberedRows()
  const out: Element[] = []
  let i = 0
  for (let r = span[0]; r <= span[1]; r++) {
    while (i < rows.length && rows[i].r < r) {
      i++
    }

    if (i < rows.length && rows[i].r === r) {
      out.push(rows[i].el)
    } else if (create) {
      const el = xml.setAttributes(xml.named(sheet.doc, 'row', X), { r: r + 1 })
      data.insertBefore(el, rows[i]?.el ?? null)
      out.push(el)
    }
  }

  return out
}

function hideRows(sheet: Sheet, span: [number, number], hidden: boolean): number {
  const rows = rowRange(sheet, span, hidden)
  for (const row of rows) {
    if (hidden) {
      row.setAttribute('hidden', '1')
    } else {
      row.removeAttribute('hidden')
    }
  }

  return rows.length
}

function colsOf(sheet: Sheet): Element {
  return (
    xml.child(sheet.root, X, 'cols') ??
    xml.insertOrdered(sheet.root, xml.named(sheet.doc, 'cols', X), SHEET_ORDER)
  )
}

function splitAt(el: Element, n: number): void {
  const min = Number(el.getAttribute('min'))
  const max = Number(el.getAttribute('max'))
  if (n <= min || n > max) {
    return
  }

  const before = xml.clone(el)
  before.setAttribute('max', String(n - 1))
  el.parentNode?.insertBefore(before, el)
  el.setAttribute('min', String(n))
}

function carve(cols: Element, first: number, last: number): Element[] {
  for (const el of xml.children(cols, X, 'col')) {
    splitAt(el, first)
    splitAt(el, last + 1)
  }

  return xml.children(cols, X, 'col').filter((el) => {
    const min = Number(el.getAttribute('min'))
    return min >= first && min <= last
  })
}

function fillGaps(sheet: Sheet, cols: Element, span: [number, number]): void {
  const first = span[0] + 1
  const last = span[1] + 1
  const covered = carve(cols, first, last)
  let next = first
  const gaps: [number, number][] = []
  for (const el of covered) {
    const min = Number(el.getAttribute('min'))
    if (min > next) {
      gaps.push([next, min - 1])
    }

    next = Number(el.getAttribute('max')) + 1
  }

  if (next <= last) {
    gaps.push([next, last])
  }

  const format = xml.child(sheet.root, X, 'sheetFormatPr')
  const width = Number(format?.getAttribute('defaultColWidth')) || DEFAULT_WIDTH
  for (const [min, max] of gaps) {
    const el = xml.setAttributes(xml.named(sheet.doc, 'col', X), { min, max, width })
    const after = xml.children(cols, X, 'col').find((c) => Number(c.getAttribute('min')) > max)
    cols.insertBefore(el, after ?? null)
  }
}

function hideColumns(sheet: Sheet, span: [number, number], hidden: boolean): number {
  const existing = xml.child(sheet.root, X, 'cols')
  if (!hidden && !existing) {
    return 0
  }

  const cols = existing ?? colsOf(sheet)
  if (hidden) {
    fillGaps(sheet, cols, span)
  }

  for (const el of carve(cols, span[0] + 1, span[1] + 1)) {
    if (hidden) {
      el.setAttribute('hidden', '1')
    } else {
      el.removeAttribute('hidden')
    }
  }

  return span[1] - span[0] + 1
}

export function hide(sheet: Sheet, opts: unknown): Hidden {
  const o = record(opts, 'hide', '{ rows?, columns?, hidden? }')
  if (o.rows === undefined && o.columns === undefined) {
    throw new DocsError('INVALID_OPTIONS', 'give rows, columns or both, like { rows: [2, 5] }')
  }

  const hidden = o.hidden === undefined ? true : o.hidden
  if (typeof hidden !== 'boolean') {
    throw new DocsError('INVALID_OPTIONS', `hidden is true or false, got ${String(hidden)}`)
  }

  const rows = o.rows === undefined ? null : rowSpan(o.rows)
  const columns = o.columns === undefined ? null : columnSpan(o.columns)
  const out: Hidden = {
    rows: rows ? hideRows(sheet, rows, hidden) : 0,
    columns: columns ? hideColumns(sheet, columns, hidden) : 0,
    hidden
  }
  sheet.touch()

  return out
}

function splitCount(value: unknown, field: string, max: number): number {
  const n = value === undefined ? 0 : value
  if (!(typeof n === 'number' && Number.isInteger(n) && n >= 0 && n < max)) {
    throw new DocsError(
      'INVALID_VALUE',
      `${field} is how many ${field} stay frozen, a whole number from 0, got ${String(value)}`
    )
  }

  return n
}

function sheetView(sheet: Sheet): Element {
  const views =
    xml.child(sheet.root, X, 'sheetViews') ??
    xml.insertOrdered(sheet.root, xml.named(sheet.doc, 'sheetViews', X), SHEET_ORDER)
  const existing = xml.child(views, X, 'sheetView')
  if (existing) {
    return existing
  }

  return xml.append(
    views,
    xml.setAttributes(xml.named(sheet.doc, 'sheetView', X), { workbookViewId: 0 })
  )
}

function activePane(rows: number, columns: number): string {
  if (rows && columns) return 'bottomRight'
  return rows ? 'bottomLeft' : 'topRight'
}

export function freeze(sheet: Sheet, opts: unknown): Frozen {
  const o = record(opts, 'freeze', '{ rows?, columns? }')
  const rows = splitCount(o.rows, 'rows', refs.MAX_ROWS)
  const columns = splitCount(o.columns, 'columns', refs.MAX_COLUMNS)
  const view = sheetView(sheet)
  for (const el of [...xml.children(view, X, 'pane'), ...xml.children(view, X, 'selection')]) {
    xml.remove(el)
  }

  sheet.touch()
  if (!rows && !columns) {
    return { rows, columns, topLeftCell: null }
  }

  const topLeftCell = refs.cell(rows, columns)
  const pane = activePane(rows, columns)
  const v = xml.builder(sheet.doc, X)
  xml.insertOrdered(
    view,
    v('pane', {
      xSplit: columns || undefined,
      ySplit: rows || undefined,
      topLeftCell,
      activePane: pane,
      state: 'frozen'
    }),
    VIEW_ORDER
  )
  xml.insertOrdered(
    view,
    v('selection', { pane, activeCell: topLeftCell, sqref: topLeftCell }),
    VIEW_ORDER
  )

  return { rows, columns, topLeftCell }
}
