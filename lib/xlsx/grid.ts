import { Unsupported } from './values.ts'
import { parse, type Node } from './expression.ts'
import type { Area } from './refs.ts'
import type { Book } from './ranges.ts'
import type { RangeValue, Scalar } from './values.ts'

interface Row<T> {
  cols: number[]
  values: Map<number, T>
}

export class Cells<T> {
  rows: Map<number, Row<T>>
  private order: number[] | null

  constructor() {
    this.rows = new Map()
    this.order = null
  }

  set(row: number, col: number, value: T): void {
    let entry = this.rows.get(row)
    if (!entry) {
      entry = { cols: [], values: new Map() }
      this.rows.set(row, entry)
      this.order = null
    }

    if (!entry.values.has(col)) {
      entry.cols.push(col)
      if (entry.cols.length > 1 && entry.cols[entry.cols.length - 2] > col) {
        entry.cols.sort((a, b) => a - b)
      }
    }

    entry.values.set(col, value)
  }

  get(row: number, col: number): T | undefined {
    return this.rows.get(row)?.values.get(col)
  }

  each(area: Area, visit: (value: T, row: number, col: number) => void): void {
    const order = (this.order ??= [...this.rows.keys()].sort((a, b) => a - b))
    for (let i = firstAtLeast(order, area.top); i < order.length && order[i] <= area.bottom; i++) {
      const row = order[i]
      const entry = this.rows.get(row)
      if (!entry) {
        continue
      }

      for (const col of entry.cols) {
        if (col > area.right) break
        const value = entry.values.get(col)
        if (col >= area.left && value !== undefined) {
          visit(value, row, col)
        }
      }
    }
  }
}

function firstAtLeast(list: number[], target: number): number {
  let lo = 0
  let hi = list.length
  while (lo < hi) {
    const mid = (lo + hi) >> 1
    if (list[mid] < target) {
      lo = mid + 1
    } else {
      hi = mid
    }
  }

  return lo
}

export interface NameEntry {
  name: string
  scope: number | null
  text: string
}

export interface GridSheet {
  key: string
  index: number
  cells: Cells<Scalar>
}

export class Grid implements Book {
  sheet: string
  scope: number
  private sheets: Map<string, GridSheet>
  private names: NameEntry[]
  private parsed: Map<NameEntry, Node>

  constructor(sheets: GridSheet[], names: NameEntry[]) {
    this.sheets = new Map(sheets.map((s) => [s.key, s]))
    this.names = names
    this.parsed = new Map()
    this.sheet = sheets[0]?.key ?? ''
    this.scope = 0
  }

  at(sheet: GridSheet): void {
    this.sheet = sheet.key
    this.scope = sheet.index
  }

  resolveSheet(name: string): string {
    const key = name.toLowerCase()
    if (!this.sheets.has(key)) {
      throw new Unsupported(`the sheet ${name}`)
    }

    return key
  }

  cell(sheet: string, row: number, col: number): Scalar {
    return this.sheets.get(sheet)?.cells.get(row, col) ?? null
  }

  each(range: RangeValue, visit: (value: Scalar, row: number, col: number) => void): void {
    this.sheets.get(range.sheet)?.cells.each(range.area, visit)
  }

  name(name: string): Node | null {
    const wanted = name.toLowerCase()
    const matches = this.names.filter((n) => n.name.toLowerCase() === wanted)
    const entry =
      matches.find((n) => n.scope === this.scope) ?? matches.find((n) => n.scope === null)
    if (!entry) {
      return null
    }

    let node = this.parsed.get(entry)
    if (!node) {
      node = parse(entry.text)
      this.parsed.set(entry, node)
    }

    return node
  }
}
