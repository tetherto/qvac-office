import { FUNCTIONS, power } from './functions.ts'
import { scalar, type Book } from './ranges.ts'
import {
  ERRORS,
  Unsupported,
  checked,
  compare,
  isError,
  toBool,
  toNumber,
  toText
} from './values.ts'
import type { BinaryNode, CallNode, Node, RefNode } from './expression.ts'
import type { RangeValue, Scalar, Value } from './values.ts'

const MAX_NAME_DEPTH = 32
const LAZY = new Set(['IF', 'IFERROR', 'IFNA'])

interface Scope {
  book: Book
  depth: number
}

function rangeOf(node: RefNode, scope: Scope): RangeValue {
  const { book } = scope
  if (node.relative && scope.depth > 0) {
    throw new Unsupported('a defined name with relative references')
  }

  const sheet = node.sheet === null ? book.sheet : book.resolveSheet(node.sheet)

  return { kind: 'range', sheet, area: node.area }
}

function span(node: BinaryNode, scope: Scope): RangeValue {
  const { left, right } = node
  if (left.type !== 'ref' || right.type !== 'ref') {
    throw new Unsupported('a range operator between computed references')
  }

  const a = rangeOf(left, scope)
  const b = rangeOf(right, scope)
  if (a.sheet !== b.sheet) {
    throw new Unsupported('a range across sheets')
  }

  return {
    kind: 'range',
    sheet: a.sheet,
    area: {
      top: Math.min(a.area.top, b.area.top),
      left: Math.min(a.area.left, b.area.left),
      bottom: Math.max(a.area.bottom, b.area.bottom),
      right: Math.max(a.area.right, b.area.right)
    }
  }
}

function arithmetic(op: string, a: Scalar, b: Scalar): Scalar {
  const x = toNumber(a)
  if (isError(x)) {
    return x
  }

  const y = toNumber(b)
  if (isError(y)) {
    return y
  }

  if (op === '+') {
    return checked(x + y)
  }

  if (op === '-') {
    return checked(x - y)
  }

  if (op === '*') {
    return checked(x * y)
  }

  if (op === '/') {
    return y === 0 ? ERRORS.div0 : checked(x / y)
  }

  const raised = power(x, y)

  return isError(raised) ? raised : checked(raised)
}

function comparison(op: string, a: Scalar, b: Scalar): Scalar {
  if (isError(a)) {
    return a
  }

  if (isError(b)) {
    return b
  }

  const order = compare(a, b)
  if (op === '=') {
    return order === 0
  }

  if (op === '<>') {
    return order !== 0
  }

  if (op === '<') {
    return order < 0
  }

  if (op === '>') {
    return order > 0
  }

  if (op === '<=') {
    return order <= 0
  }

  return order >= 0
}

function concatenation(a: Scalar, b: Scalar): Scalar {
  const x = toText(a)
  if (isError(x)) {
    return x
  }

  const y = toText(b)
  if (isError(y)) {
    return y
  }

  return x.length + y.length > 32767 ? ERRORS.value : x + y
}

function binary(node: BinaryNode, scope: Scope): Value {
  if (node.op === ':') {
    return span(node, scope)
  }

  const a = scalar(evaluateIn(node.left, scope), scope.book)
  const b = scalar(evaluateIn(node.right, scope), scope.book)
  if (node.op === '&') {
    return concatenation(a, b)
  }

  if (['+', '-', '*', '/', '^'].includes(node.op)) {
    return arithmetic(node.op, a, b)
  }

  return comparison(node.op, a, b)
}

function lazy(node: CallNode, scope: Scope): Value {
  const { name, args } = node
  const fewest = name === 'IF' ? 1 : 2
  const most = name === 'IF' ? 3 : 2
  if (args.length < fewest || args.length > most) {
    throw new Unsupported(`${name} with ${args.length} arguments`)
  }

  const first = scalar(evaluateIn(args[0], scope), scope.book)
  if (name !== 'IF') {
    const caught = isError(first) && (name === 'IFERROR' || first === ERRORS.na)

    return caught ? evaluateIn(args[1], scope) : first
  }

  const test = toBool(first)
  if (isError(test)) {
    return test
  }

  const branch = test ? args[1] : args[2]
  if (branch === undefined) {
    return test ? 0 : false
  }

  return evaluateIn(branch, scope)
}

function call(node: CallNode, scope: Scope): Value {
  if (LAZY.has(node.name)) {
    return lazy(node, scope)
  }

  const fn = FUNCTIONS[node.name]
  if (!fn) {
    throw new Unsupported(`the function ${node.name}`)
  }

  return fn(
    node.args.map((arg) => evaluateIn(arg, scope)),
    scope.book
  )
}

function named(name: string, scope: Scope): Node {
  const node = scope.book.name(name)
  if (!node || scope.depth >= MAX_NAME_DEPTH) {
    throw new Unsupported(`the name ${name}`)
  }

  return node
}

function deeper(scope: Scope): Scope {
  return { book: scope.book, depth: scope.depth + 1 }
}

function evaluateIn(node: Node, scope: Scope): Value {
  switch (node.type) {
    case 'literal':
      return node.value
    case 'missing':
      return null
    case 'array':
      return { kind: 'array', rows: node.rows }
    case 'ref':
      return rangeOf(node, scope)
    case 'name':
      return evaluateIn(named(node.name, scope), deeper(scope))
    case 'unary': {
      const n = toNumber(scalar(evaluateIn(node.arg, scope), scope.book))

      return isError(n) || node.op === '+' ? n : -n
    }
    case 'percent': {
      const n = toNumber(scalar(evaluateIn(node.arg, scope), scope.book))

      return isError(n) ? n : n / 100
    }
    case 'binary':
      return binary(node, scope)
    case 'call':
      return call(node, scope)
  }
}

export function evaluate(node: Node, book: Book): Value {
  return evaluateIn(node, { book, depth: 0 })
}

export function result(node: Node, book: Book): Scalar {
  const value = scalar(evaluate(node, book), book)
  if (value === null) {
    return 0
  }

  if (typeof value === 'number' && Object.is(value, -0)) {
    return 0
  }

  return value
}

function children(node: Node): Node[] {
  if (node.type === 'call') {
    return node.args
  }

  if (node.type === 'unary' || node.type === 'percent') {
    return [node.arg]
  }

  if (node.type === 'binary') {
    return [node.left, node.right]
  }

  return []
}

function collect(node: Node, scope: Scope, out: RangeValue[]): void {
  if (node.type === 'ref') {
    out.push(rangeOf(node, scope))
  } else if (node.type === 'name') {
    collect(named(node.name, scope), deeper(scope), out)
  } else if (node.type === 'binary' && node.op === ':') {
    out.push(span(node, scope))
  } else if (node.type === 'call' && !LAZY.has(node.name) && !FUNCTIONS[node.name]) {
    throw new Unsupported(`the function ${node.name}`)
  }

  for (const child of children(node)) {
    collect(child, scope, out)
  }
}

export function dependencies(node: Node, book: Book): RangeValue[] {
  const out: RangeValue[] = []
  collect(node, { book, depth: 0 }, out)

  return out
}
