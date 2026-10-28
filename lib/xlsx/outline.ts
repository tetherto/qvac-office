import * as refs from './refs.ts'
import * as tables from './tables.ts'
import type { Area } from './refs.ts'
import type { Element } from '../xml.ts'
import type { ReshapeHost } from './reshape.ts'
import type { CellValue, Sheet, SheetEntry } from './sheet.ts'

export interface HeaderGuess {
  row: number
  names: string[]
}

export interface SheetOutline {
  index: number
  name: string
  kind: string
  state: string
  range: string | null
  header: HeaderGuess | null
  rows: number
  columns: number
  charts: number
  tables: number
}

interface Grid {
  sheet: Sheet
  area: Area
}

const CHART_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
const HEADER_SCAN = 10
const MAX_NAMES = 20
const MAX_NAME_LENGTH = 40

function rowValues(grid: Grid, row: Element, r: number): (CellValue | undefined)[] {
  const { sheet, area } = grid
  const values: (CellValue | undefined)[] = new Array(area.right - area.left + 1).fill(undefined)
  for (const { el, col } of sheet.numberedCells(row, r)) {
    if (col < area.left || col > area.right) continue
    values[col - area.left] = sheet.formula(el) ? 0 : sheet.value(el)
  }

  return values
}

function isHeader(values: (CellValue | undefined)[]): boolean {
  const filled = values.filter((v) => v !== undefined && v !== null && v !== '')
  return (
    filled.length >= Math.max(Math.min(2, values.length), Math.ceil(values.length / 2)) &&
    filled.every((v) => typeof v === 'string')
  )
}

function guessHeader(grid: Grid): HeaderGuess | null {
  const rows = grid.sheet
    .numberedRows()
    .filter(({ r }) => r >= grid.area.top && r < grid.area.top + HEADER_SCAN)
  for (const { el, r } of rows) {
    const values = rowValues(grid, el, r)
    if (!isHeader(values)) {
      continue
    }

    const names = values
      .slice(0, MAX_NAMES)
      .map((v) => (typeof v === 'string' ? v.slice(0, MAX_NAME_LENGTH) : ''))

    return { row: r + 1, names }
  }

  return null
}

function chartCount(host: ReshapeHost, sheet: Sheet): number {
  const drawing = host.drawingOf(sheet)
  return drawing ? host.package.related(drawing, CHART_REL).length : 0
}

function outlineOf(host: ReshapeHost, entry: SheetEntry): SheetOutline {
  const sheet = host.sheetFor(entry)
  const out: SheetOutline = {
    index: entry.index,
    name: entry.name,
    kind: entry.kind,
    state: entry.state,
    range: null,
    header: null,
    rows: 0,
    columns: 0,
    charts: chartCount(host, sheet),
    tables: 0
  }
  if (entry.kind !== 'worksheet') {
    return out
  }

  const area = sheet.usedRange()
  out.tables = tables.on(sheet).length
  if (area) {
    out.range = refs.range(area)
    out.rows = area.bottom - area.top + 1
    out.columns = area.right - area.left + 1
    out.header = guessHeader({ sheet, area })
  }

  return out
}

export function outline(host: ReshapeHost): SheetOutline[] {
  return host.entries().map((entry) => outlineOf(host, entry))
}
