import { Unsupported, isArray, isRange } from './values.ts'
import type { ArrayValue, RangeValue, Scalar, Value } from './values.ts'
import type { Node } from './expression.ts'
import type { Position } from './refs.ts'

export interface Book {
  sheet: string
  resolveSheet(name: string): string
  cell(sheet: string, row: number, col: number): Scalar
  each(range: RangeValue, visit: (value: Scalar, row: number, col: number) => void): void
  name(name: string): Node | null
}

export type Block = RangeValue | ArrayValue

export interface Size {
  rows: number
  cols: number
}

export function isBlock(value: Value): value is Block {
  return isRange(value) || isArray(value)
}

export function size(block: Block): Size {
  if (isArray(block)) {
    return { rows: block.rows.length, cols: block.rows[0].length }
  }

  const { area } = block

  return { rows: area.bottom - area.top + 1, cols: area.right - area.left + 1 }
}

export function at(block: Block, book: Book, pos: Position): Scalar {
  const { row, col } = pos
  if (isArray(block)) {
    return block.rows[row][col]
  }

  return book.cell(block.sheet, block.area.top + row, block.area.left + col)
}

export function populated(
  block: Block,
  book: Book,
  visit: (value: Scalar, row: number, col: number) => void
): void {
  if (isArray(block)) {
    block.rows.forEach((values, r) => values.forEach((value, c) => visit(value, r, c)))
    return
  }

  const { top, left } = block.area
  book.each(block, (value, row, col) => visit(value, row - top, col - left))
}

export function scalar(value: Value, book: Book): Scalar {
  if (!isBlock(value)) {
    return value
  }

  const { rows, cols } = size(value)
  if (rows !== 1 || cols !== 1) {
    throw new Unsupported('a range where one value is expected')
  }

  return at(value, book, { row: 0, col: 0 })
}

export function block(value: Value): Block {
  if (!isBlock(value)) {
    throw new Unsupported('a value where a range is expected')
  }

  return value
}

export function sub(source: Block, from: Size, extent: Size): Block {
  if (isArray(source)) {
    const rows = source.rows
      .slice(from.rows, from.rows + extent.rows)
      .map((row) => row.slice(from.cols, from.cols + extent.cols))
    return { kind: 'array', rows }
  }

  const top = source.area.top + from.rows
  const left = source.area.left + from.cols

  return {
    kind: 'range',
    sheet: source.sheet,
    area: { top, left, bottom: top + extent.rows - 1, right: left + extent.cols - 1 }
  }
}
