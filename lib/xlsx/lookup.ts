import { ERRORS, Unsupported, compare, isError, toBool, toNumber } from './values.ts'
import { wildcard } from './criteria.ts'
import { at, block, populated, scalar, size, sub, type Block, type Book } from './ranges.ts'
import type { ErrorValue, Scalar, Value } from './values.ts'
import type { Fn } from './functions.ts'

interface Entry {
  index: number
  value: Scalar
}

interface Search {
  vector: Block
  key: Scalar
  mode: number
}

function entries(vector: Block, book: Book): Entry[] {
  const { rows } = size(vector)
  const out: Entry[] = []
  populated(vector, book, (value, r, c) => {
    if (value !== null) {
      out.push({ index: rows === 1 ? c : r, value })
    }
  })

  return out.sort((a, b) => a.index - b.index)
}

function sameValue(key: Scalar): (value: Scalar) => boolean {
  if (typeof key === 'string') {
    const re = wildcard(key)
    return (value) => typeof value === 'string' && re.test(value)
  }

  return (value) => typeof value === typeof key && compare(value, key) === 0
}

function sorted(list: Entry[], descending: boolean): boolean {
  for (let i = 1; i < list.length; i++) {
    const order = compare(list[i - 1].value, list[i].value)
    if (descending ? order < 0 : order > 0) {
      return false
    }
  }

  return true
}

function contiguous(list: Entry[]): boolean {
  return list.every((entry, i) => entry.index === (list[0]?.index ?? 0) + i)
}

function search(spec: Search, book: Book): number | ErrorValue {
  const { vector, key, mode } = spec
  if (key === null) {
    throw new Unsupported('a lookup of a blank cell')
  }

  if (isError(key)) {
    return key
  }

  const list = entries(vector, book)
  if (mode === 0) {
    const same = sameValue(key)
    return list.find((entry) => same(entry.value))?.index ?? ERRORS.na
  }

  if (!sorted(list, mode < 0) || !contiguous(list) || (list.length && list[0].index !== 0)) {
    throw new Unsupported('an approximate lookup over unsorted or gapped data')
  }

  let found = -1
  for (const entry of list.filter((e) => typeof e.value === typeof key)) {
    const order = compare(entry.value, key)
    if (mode > 0 ? order <= 0 : order >= 0) {
      found = entry.index
    } else {
      break
    }
  }

  return found < 0 ? ERRORS.na : found
}

function approximate(args: Value[], i: number, book: Book): boolean | ErrorValue {
  return args.length > i ? toBool(scalar(args[i], book)) : true
}

function tableLookup(vertical: boolean): Fn {
  return (args, book) => {
    if (args.length < 3 || args.length > 4) {
      throw new Unsupported('a lookup with the wrong arguments')
    }

    const table = block(args[1])
    const offset = toNumber(scalar(args[2], book))
    const approx = approximate(args, 3, book)
    if (isError(offset)) {
      return offset
    }

    if (isError(approx)) {
      return approx
    }

    const n = Math.trunc(offset)
    const { rows, cols } = size(table)
    if (n < 1) {
      return ERRORS.value
    }

    if (n > (vertical ? cols : rows)) {
      return ERRORS.ref
    }

    const vector = vertical
      ? sub(table, { rows: 0, cols: 0 }, { rows, cols: 1 })
      : sub(table, { rows: 0, cols: 0 }, { rows: 1, cols })
    const found = search({ vector, key: scalar(args[0], book), mode: approx ? 1 : 0 }, book)
    if (isError(found)) {
      return found
    }

    const pos = vertical ? { row: found, col: n - 1 } : { row: n - 1, col: found }

    return at(table, book, pos)
  }
}

function match(args: Value[], book: Book): Value {
  if (args.length < 2 || args.length > 3) {
    throw new Unsupported('MATCH with the wrong arguments')
  }

  const vector = block(args[1])
  const { rows, cols } = size(vector)
  const mode = args.length > 2 ? toNumber(scalar(args[2], book)) : 1
  if (isError(mode)) {
    return mode
  }

  if (rows !== 1 && cols !== 1) {
    return ERRORS.na
  }

  const found = search({ vector, key: scalar(args[0], book), mode: Math.sign(mode) }, book)

  return isError(found) ? found : found + 1
}

function indexArg(args: Value[], i: number, book: Book): number | ErrorValue {
  if (args.length <= i) return 0
  const n = toNumber(scalar(args[i], book))
  if (isError(n)) {
    return n
  }

  return n < 0 ? ERRORS.value : Math.trunc(n)
}

function index(args: Value[], book: Book): Value {
  if (args.length < 2 || args.length > 3) {
    throw new Unsupported('INDEX with the wrong arguments')
  }

  const source = block(args[0])
  const { rows, cols } = size(source)
  const first = indexArg(args, 1, book)
  const second = indexArg(args, 2, book)
  if (isError(first)) {
    return first
  }

  if (isError(second)) {
    return second
  }

  const [r, c] = rows === 1 && args.length === 2 ? [1, first] : [first, second]
  if (r > rows || c > cols) {
    return ERRORS.ref
  }

  const from = { rows: r === 0 ? 0 : r - 1, cols: c === 0 ? 0 : c - 1 }
  const extent = { rows: r === 0 ? rows : 1, cols: c === 0 ? cols : 1 }

  return sub(source, from, extent)
}

export const lookups: Record<string, Fn> = {
  VLOOKUP: tableLookup(true),
  HLOOKUP: tableLookup(false),
  MATCH: match,
  INDEX: index
}
