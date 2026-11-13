import {
  ERRORS,
  Unsupported,
  checked,
  isError,
  rounded,
  textToNumber,
  toBool,
  toNumber,
  toText
} from './values.ts'
import { isBlock, populated, scalar, size, type Book } from './ranges.ts'
import { conditional } from './aggregate-if.ts'
import { lookups } from './lookup.ts'
import type { ErrorValue, Scalar, Value } from './values.ts'

export type Fn = (args: Value[], book: Book) => Value

interface Mode {
  strict: boolean
}

interface TextMode {
  keepEmpty: boolean
}

interface Arity {
  min: number
  max: number
}

const MAX_TEXT = 32767

function arity(name: string, args: Value[], range: Arity): void {
  if (args.length < range.min || args.length > range.max) {
    throw new Unsupported(`${name} with ${args.length} arguments`)
  }
}

function numberArg(value: Value, book: Book): number | ErrorValue {
  return toNumber(scalar(value, book))
}

function textArg(value: Value, book: Book): string | ErrorValue {
  return toText(scalar(value, book))
}

function directNumber(value: Scalar, strict: boolean): number | ErrorValue | null {
  if (typeof value === 'string') {
    const n = textToNumber(value)
    if (n !== null) {
      return n
    }

    return strict ? ERRORS.value : null
  }

  if (isError(value)) {
    return strict ? value : null
  }

  return toNumber(value)
}

function collect(args: Value[], book: Book, strict: boolean): number[] | ErrorValue {
  const out: number[] = []
  let failed: ErrorValue | null = null
  for (const arg of args) {
    if (isBlock(arg)) {
      populated(arg, book, (value) => {
        if (typeof value === 'number') {
          out.push(value)
        } else if (strict && isError(value)) {
          failed ??= value
        }
      })
    } else {
      const n = directNumber(arg, strict)
      if (isError(n)) {
        failed ??= n
      } else if (n !== null) {
        out.push(n)
      }
    }

    if (failed) {
      return failed
    }
  }

  return out
}

function aggregate(reduce: (numbers: number[]) => number | ErrorValue, mode: Mode): Fn {
  const { strict } = mode
  return (args, book) => {
    const numbers = collect(args, book, strict)
    return isError(numbers) ? numbers : reduce(numbers)
  }
}

function counta(args: Value[], book: Book): number {
  let count = 0
  for (const arg of args) {
    if (isBlock(arg)) {
      populated(arg, book, (value) => {
        if (value !== null) {
          count++
        }
      })
    } else {
      count++
    }
  }

  return count
}

function countblank(args: Value[], book: Book): number {
  arity('COUNTBLANK', args, { min: 1, max: 1 })
  const range = args[0]
  if (!isBlock(range)) {
    throw new Unsupported('COUNTBLANK of a value')
  }

  const { rows, cols } = size(range)
  let filled = 0
  populated(range, book, (value) => {
    if (value !== null && value !== '') {
      filled++
    }
  })

  return rows * cols - filled
}

function logical(args: Value[], book: Book): boolean[] | ErrorValue {
  const out: boolean[] = []
  for (const arg of args) {
    if (isBlock(arg)) {
      let failed: ErrorValue | null = null
      populated(arg, book, (value) => {
        if (typeof value === 'boolean' || typeof value === 'number') {
          out.push(toBool(value) === true)
        } else if (isError(value)) {
          failed ??= value
        }
      })
      if (failed) {
        return failed
      }

      continue
    }

    const b = toBool(arg)
    if (isError(b)) {
      return b
    }

    out.push(b)
  }

  return out.length ? out : ERRORS.value
}

function andOr(every: boolean): Fn {
  return (args, book) => {
    const values = logical(args, book)
    if (isError(values)) {
      return values
    }

    return every ? values.every(Boolean) : values.some(Boolean)
  }
}

function unaryMath(fn: (n: number) => number | ErrorValue): Fn {
  return (args, book) => {
    arity('a math function', args, { min: 1, max: 1 })
    const n = numberArg(args[0], book)
    if (isError(n)) {
      return n
    }

    const out = fn(n)

    return isError(out) ? out : checked(out)
  }
}

function binaryMath(fn: (a: number, b: number) => number | ErrorValue, fallback?: number): Fn {
  return (args, book) => {
    arity('a math function', args, { min: fallback === undefined ? 2 : 1, max: 2 })
    const a = numberArg(args[0], book)
    if (isError(a)) {
      return a
    }

    const b = args.length > 1 ? numberArg(args[1], book) : (fallback ?? 0)
    if (isError(b)) {
      return b
    }

    const out = fn(a, b)

    return isError(out) ? out : checked(out)
  }
}

function scaled(n: number, digits: number, round: (x: number) => number): number {
  const d = Math.trunc(digits)
  const factor = 10 ** Math.abs(d)
  const magnitude = rounded(d >= 0 ? Math.abs(n) * factor : Math.abs(n) / factor)
  const whole = round(magnitude)
  const back = d >= 0 ? whole / factor : whole * factor

  return Math.sign(n) * rounded(back)
}

export function power(a: number, b: number): number | ErrorValue {
  if (a === 0 && b === 0) return ERRORS.num
  if (a === 0 && b < 0) return ERRORS.div0
  if (a < 0 && !Number.isInteger(b)) {
    return ERRORS.num
  }

  return a ** b
}

function mod(n: number, d: number): number | ErrorValue {
  if (d === 0) return ERRORS.div0
  return rounded(n - d * Math.floor(n / d))
}

const math: Record<string, Fn> = {
  ROUND: binaryMath((n, d) => scaled(n, d, (x) => Math.floor(x + 0.5))),
  ROUNDUP: binaryMath((n, d) => scaled(n, d, Math.ceil)),
  ROUNDDOWN: binaryMath((n, d) => scaled(n, d, Math.floor)),
  ABS: unaryMath(Math.abs),
  INT: unaryMath(Math.floor),
  MOD: binaryMath(mod),
  POWER: binaryMath(power),
  SQRT: unaryMath((n) => (n < 0 ? ERRORS.num : Math.sqrt(n)))
}

function joined(parts: string[]): string | ErrorValue {
  const out = parts.join('')
  return out.length > MAX_TEXT ? ERRORS.value : out
}

function texts(args: Value[], book: Book, mode: TextMode): string[] | ErrorValue {
  const { keepEmpty } = mode
  const out: string[] = []
  for (const arg of args) {
    let failed: ErrorValue | null = null
    const add = (value: Scalar): void => {
      const text = toText(value)
      if (isError(text)) {
        failed ??= text
      } else if (keepEmpty || text !== '') {
        out.push(text)
      }
    }
    if (isBlock(arg) && keepEmpty) {
      everyCell(arg, book, add)
    } else if (isBlock(arg)) {
      populated(arg, book, add)
    } else {
      add(arg)
    }

    if (failed) {
      return failed
    }
  }

  return out
}

function everyCell(arg: Value, book: Book, add: (value: Scalar) => void): void {
  if (!isBlock(arg)) return
  const { rows, cols } = size(arg)
  if (rows * cols > 100000) {
    throw new Unsupported('TEXTJOIN over a very large range')
  }

  const grid: Scalar[][] = Array.from({ length: rows }, () => new Array(cols).fill(null))
  populated(arg, book, (value, r, c) => {
    grid[r][c] = value
  })
  grid.flat().forEach(add)
}

function concat(args: Value[], book: Book): Value {
  const parts = texts(args, book, { keepEmpty: false })
  return isError(parts) ? parts : joined(parts)
}

function concatenate(args: Value[], book: Book): Value {
  const parts = texts(
    args.map((arg) => scalar(arg, book)),
    book,
    { keepEmpty: false }
  )
  return isError(parts) ? parts : joined(parts)
}

function textjoin(args: Value[], book: Book): Value {
  arity('TEXTJOIN', args, { min: 3, max: 254 })
  const delimiter = textArg(args[0], book)
  if (isError(delimiter)) {
    return delimiter
  }

  const ignore = toBool(scalar(args[1], book))
  if (isError(ignore)) {
    return ignore
  }

  const parts = texts(args.slice(2), book, { keepEmpty: true })
  if (isError(parts)) {
    return parts
  }

  const kept = ignore ? parts.filter((p) => p !== '') : parts
  const out = kept.join(delimiter)

  return out.length > MAX_TEXT ? ERRORS.value : out
}

function textFn(fn: (text: string) => Value): Fn {
  return (args, book) => {
    arity('a text function', args, { min: 1, max: 1 })
    const text = textArg(args[0], book)
    return isError(text) ? text : fn(text)
  }
}

function countArg(args: Value[], i: number, book: Book): number | ErrorValue {
  if (args.length <= i) return 1
  const n = numberArg(args[i], book)
  if (isError(n)) {
    return n
  }

  return n < 0 ? ERRORS.value : Math.trunc(n)
}

function side(fromEnd: boolean): Fn {
  return (args, book) => {
    arity('LEFT or RIGHT', args, { min: 1, max: 2 })
    const text = textArg(args[0], book)
    if (isError(text)) {
      return text
    }

    const n = countArg(args, 1, book)
    if (isError(n)) {
      return n
    }

    if (fromEnd) {
      return n === 0 ? '' : text.slice(-n)
    }

    return text.slice(0, n)
  }
}

function mid(args: Value[], book: Book): Value {
  arity('MID', args, { min: 3, max: 3 })
  const text = textArg(args[0], book)
  if (isError(text)) {
    return text
  }

  const start = numberArg(args[1], book)
  if (isError(start)) {
    return start
  }

  const n = countArg(args, 2, book)
  if (isError(n)) {
    return n
  }

  if (start < 1) {
    return ERRORS.value
  }

  const from = Math.trunc(start) - 1

  return text.slice(from, from + n)
}

const text: Record<string, Fn> = {
  CONCAT: concat,
  CONCATENATE: concatenate,
  TEXTJOIN: textjoin,
  LEN: textFn((t) => t.length),
  LEFT: side(false),
  RIGHT: side(true),
  MID: mid,
  UPPER: textFn((t) => t.toUpperCase()),
  LOWER: textFn((t) => t.toLowerCase()),
  TRIM: textFn((t) => t.replace(/ +/g, ' ').replace(/^ | $/g, ''))
}

function sum(numbers: number[]): number {
  return numbers.reduce((a, b) => a + b, 0)
}

const aggregates: Record<string, Fn> = {
  SUM: aggregate(sum, { strict: true }),
  AVERAGE: aggregate((n) => (n.length ? sum(n) / n.length : ERRORS.div0), { strict: true }),
  MIN: aggregate((n) => (n.length ? n.reduce((a, b) => Math.min(a, b)) : 0), { strict: true }),
  MAX: aggregate((n) => (n.length ? n.reduce((a, b) => Math.max(a, b)) : 0), { strict: true }),
  COUNT: aggregate((n) => n.length, { strict: false }),
  COUNTA: counta,
  COUNTBLANK: countblank
}

const logic: Record<string, Fn> = {
  TRUE: () => true,
  FALSE: () => false,
  AND: andOr(true),
  OR: andOr(false),
  NOT: (args, book) => {
    arity('NOT', args, { min: 1, max: 1 })
    const b = toBool(scalar(args[0], book))
    return isError(b) ? b : !b
  }
}

export const FUNCTIONS: Record<string, Fn> = {
  ...aggregates,
  ...logic,
  ...math,
  ...text,
  ...conditional,
  ...lookups
}
