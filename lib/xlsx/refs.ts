import { DocsError } from '../errors.ts'

export interface Position {
  row: number
  col: number
}

export interface Area {
  top: number
  left: number
  bottom: number
  right: number
}

export const MAX_ROWS = 1048576
export const MAX_COLUMNS = 16384

export function column(index: number): string {
  let n = index + 1
  let out = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    out = String.fromCharCode(65 + rem) + out
    n = Math.floor((n - 1) / 26)
  }

  return out
}

export function columnIndex(letters: string): number {
  let n = 0
  for (const ch of letters.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64)
  }

  return n - 1
}

export function cell(row: number, col: number): string {
  return `${column(col)}${row + 1}`
}

export function parse(ref: unknown): Position {
  const match =
    typeof ref === 'string' ? /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/.exec(ref.trim()) : null
  if (!match) {
    throw new DocsError('INVALID_CELL', `"${String(ref)}" is not a cell reference like A1 or C12`)
  }

  const col = columnIndex(match[1])
  const row = Number(match[2]) - 1
  if (row < 0 || row >= MAX_ROWS || col >= MAX_COLUMNS) {
    throw new DocsError('INVALID_CELL', `${String(ref)} is outside the sheet`)
  }

  return { row, col }
}

export function parseRange(ref: unknown): Area {
  if (typeof ref !== 'string') {
    throw new DocsError('INVALID_RANGE', 'a range is text like A1:D20')
  }

  const [a, b = a] = ref.split(':')
  const start = parse(a)
  const end = parse(b)

  return {
    top: Math.min(start.row, end.row),
    left: Math.min(start.col, end.col),
    bottom: Math.max(start.row, end.row),
    right: Math.max(start.col, end.col)
  }
}

export function range(area: Area): string {
  const start = cell(area.top, area.left)
  return area.top === area.bottom && area.left === area.right
    ? start
    : `${start}:${cell(area.bottom, area.right)}`
}

export function quoteSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}[0-9]+$/.test(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`
}

export function absolute(sheet: string, area: Area): string {
  const start = `$${column(area.left)}$${area.top + 1}`
  if (area.top === area.bottom && area.left === area.right) {
    return `${quoteSheet(sheet)}!${start}`
  }

  return `${quoteSheet(sheet)}!${start}:$${column(area.right)}$${area.bottom + 1}`
}

export function columnArg(value: unknown): number {
  if (typeof value === 'string' && /^[A-Za-z]{1,3}$/.test(value)) {
    return columnIndex(value)
  }

  if (typeof value === 'number' && Number.isInteger(value) && value >= 1 && value <= MAX_COLUMNS) {
    return value - 1
  }

  throw new DocsError(
    'INVALID_CELL',
    `a column is a letter like "C" or a number from 1, got ${String(value)}`
  )
}

export function rowArg(value: unknown, name = 'row'): number {
  if (typeof value === 'number' && Number.isInteger(value) && value >= 1 && value <= MAX_ROWS) {
    return value - 1
  }

  throw new DocsError('INVALID_CELL', `${name} is a row number from 1, got ${String(value)}`)
}
