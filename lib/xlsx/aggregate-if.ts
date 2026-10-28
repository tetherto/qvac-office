import { ERRORS, Unsupported, isError, isRange } from './values.ts'
import { criterion, type Criterion } from './criteria.ts'
import { at, populated, scalar, size, type Book } from './ranges.ts'
import type { ErrorValue, RangeValue, Scalar, Value } from './values.ts'
import type { Fn } from './functions.ts'

interface Tally {
  sum: number
  count: number
  error: ErrorValue | null
}

interface Plan {
  range: RangeValue
  test: Criterion
  target: RangeValue
}

function rangeArg(value: Value): RangeValue {
  if (!isRange(value)) {
    throw new Unsupported('a criteria function over a value')
  }

  return value
}

function plan(args: Value[], book: Book): Plan {
  if (args.length < 2 || args.length > 3) {
    throw new Unsupported('a criteria function with the wrong arguments')
  }

  const range = rangeArg(args[0])
  const test = criterion(scalar(args[1], book))
  const given = args.length === 3 ? rangeArg(args[2]) : range
  const { rows, cols } = size(range)
  const { top, left } = given.area
  const target: RangeValue = {
    kind: 'range',
    sheet: given.sheet,
    area: { top, left, bottom: top + rows - 1, right: left + cols - 1 }
  }

  return { range, test, target }
}

function add(tally: Tally, value: Scalar): void {
  if (typeof value === 'number') {
    tally.sum += value
    tally.count++
  } else if (isError(value)) {
    tally.error ??= value
  }
}

function tally(args: Value[], book: Book): Tally {
  const { range, test, target } = plan(args, book)
  const out: Tally = { sum: 0, count: 0, error: null }
  const matched = new Set<string>()
  populated(range, book, (value, r, c) => {
    if (value === null ? !test.blank : !test.test(value)) return
    matched.add(`${r}:${c}`)
    add(out, at(target, book, { row: r, col: c }))
  })

  if (test.blank) {
    populated(target, book, (value, r, c) => {
      if (matched.has(`${r}:${c}`) || at(range, book, { row: r, col: c }) !== null) {
        return
      }

      add(out, value)
    })
  }

  return out
}

function countif(args: Value[], book: Book): Value {
  if (args.length !== 2) {
    throw new Unsupported('COUNTIF with the wrong arguments')
  }

  const range = rangeArg(args[0])
  const test = criterion(scalar(args[1], book))
  const { rows, cols } = size(range)
  let count = 0
  let filled = 0
  populated(range, book, (value) => {
    if (value === null) return
    filled++
    if (test.test(value)) {
      count++
    }
  })

  return test.blank ? count + rows * cols - filled : count
}

function sumif(args: Value[], book: Book): Value {
  const out = tally(args, book)
  return out.error ?? out.sum
}

function averageif(args: Value[], book: Book): Value {
  const out = tally(args, book)
  if (out.error) {
    return out.error
  }

  return out.count ? out.sum / out.count : ERRORS.div0
}

export const conditional: Record<string, Fn> = {
  SUMIF: sumif,
  COUNTIF: countif,
  AVERAGEIF: averageif
}
