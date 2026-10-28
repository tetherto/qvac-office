import * as refs from './refs.ts'
import { Unsupported, errorNamed, type Scalar } from './values.ts'
import type { Area, Position } from './refs.ts'

export interface RefNode {
  type: 'ref'
  sheet: string | null
  area: Area
  relative: boolean
}

export interface CallNode {
  type: 'call'
  name: string
  args: Node[]
}

export interface BinaryNode {
  type: 'binary'
  op: string
  left: Node
  right: Node
}

export type Node =
  | { type: 'literal'; value: Scalar }
  | RefNode
  | { type: 'name'; name: string }
  | CallNode
  | { type: 'unary'; op: string; arg: Node }
  | { type: 'percent'; arg: Node }
  | BinaryNode
  | { type: 'array'; rows: Scalar[][] }
  | { type: 'missing' }

type Token =
  | { kind: 'literal'; value: Scalar }
  | { kind: 'ref'; node: RefNode }
  | { kind: 'name'; text: string }
  | { kind: 'func'; text: string }
  | { kind: 'op'; text: string }

const ERROR_LITERAL = /#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!)/iy
const NUMBER = /(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?/y
const CELL =
  /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?(?![A-Za-z0-9_.(![])/y
const COLUMNS = /(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_.(![])/y
const ROWS = /(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![A-Za-z0-9_.(![])/y
const IDENTIFIER = /[A-Za-z_\\À-￿][\w.?\\À-￿]*/y
const BARE_SHEET = /[A-Za-z_À-￿][\w.À-￿]*!/y
const OPERATORS = ['<>', '<=', '>=', '+', '-', '*', '/', '^', '&', '=', '<', '>', '%', '(', ')']
const PUNCTUATION = [',', ';', '{', '}', ':']

function matchAt(re: RegExp, text: string, at: number): RegExpExecArray | null {
  re.lastIndex = at
  return re.exec(text)
}

function area(from: Position, to: Position): Area {
  const [top, left, bottom, right] = [from.row, from.col, to.row, to.col]

  return {
    top: Math.min(top, bottom),
    left: Math.min(left, right),
    bottom: Math.max(top, bottom),
    right: Math.max(left, right)
  }
}

function inSheet(col: string, row?: string): boolean {
  const okRow = row === undefined || (Number(row) >= 1 && Number(row) <= refs.MAX_ROWS)
  return refs.columnIndex(col) < refs.MAX_COLUMNS && okRow
}

function cellRef(m: RegExpExecArray): RefNode | null {
  const end = m[6] ? [m[6], m[8]] : [m[2], m[4]]
  if (!inSheet(m[2], m[4]) || !inSheet(end[0], end[1])) {
    return null
  }

  const relative =
    m[1] !== '$' || m[3] !== '$' || (m[6] !== undefined && (m[5] !== '$' || m[7] !== '$'))

  return {
    type: 'ref',
    sheet: null,
    relative,
    area: area(
      { row: Number(m[4]) - 1, col: refs.columnIndex(m[2]) },
      { row: Number(end[1]) - 1, col: refs.columnIndex(end[0]) }
    )
  }
}

function wholeRef(text: string, at: number): [RefNode, number] | null {
  const cols = matchAt(COLUMNS, text, at)
  if (cols && inSheet(cols[2]) && inSheet(cols[4])) {
    const relative = cols[1] !== '$' || cols[3] !== '$'
    const box = area(
      { row: 0, col: refs.columnIndex(cols[2]) },
      { row: refs.MAX_ROWS - 1, col: refs.columnIndex(cols[4]) }
    )

    return [{ type: 'ref', sheet: null, relative, area: box }, cols[0].length]
  }

  const rows = matchAt(ROWS, text, at)
  if (rows && inSheet('A', rows[2]) && inSheet('A', rows[4])) {
    const relative = rows[1] !== '$' || rows[3] !== '$'
    const box = area(
      { row: Number(rows[2]) - 1, col: 0 },
      { row: Number(rows[4]) - 1, col: refs.MAX_COLUMNS - 1 }
    )

    return [{ type: 'ref', sheet: null, relative, area: box }, rows[0].length]
  }

  return null
}

function readRef(text: string, at: number): [RefNode, number] | null {
  const cell = matchAt(CELL, text, at)
  const node = cell ? cellRef(cell) : null
  if (cell && node) {
    return [node, cell[0].length]
  }

  return wholeRef(text, at)
}

function quotedEnd(text: string, at: number, quote: string): number {
  let i = at + 1
  while (i < text.length) {
    if (text[i] === quote && text[i + 1] === quote) {
      i += 2
    } else if (text[i] === quote) {
      return i + 1
    } else {
      i++
    }
  }

  throw new Unsupported('an unclosed quote')
}

interface Lexer {
  text: string
  at: number
  out: Token[]
}

function sheetPrefix(lex: Lexer): string | null {
  const { text, at } = lex
  if (text[at] === "'") {
    const end = quotedEnd(text, at, "'")
    if (text[end] !== '!') {
      throw new Unsupported('a quoted name that is not a sheet')
    }

    lex.at = end + 1

    return text.slice(at + 1, end - 1).replace(/''/g, "'")
  }

  const bare = matchAt(BARE_SHEET, text, at)
  if (!bare) {
    return null
  }

  lex.at = at + bare[0].length

  return bare[0].slice(0, -1)
}

function sheetToken(lex: Lexer, sheet: string): void {
  if (sheet.includes(':') || sheet.includes('[')) {
    throw new Unsupported('3D and external references')
  }

  const error = matchAt(ERROR_LITERAL, lex.text, lex.at)
  if (error) {
    lex.out.push({ kind: 'literal', value: errorNamed(error[0]) })
    lex.at += error[0].length
    return
  }

  const ref = readRef(lex.text, lex.at)
  if (!ref) {
    throw new Unsupported('a sheet-scoped name')
  }

  lex.out.push({ kind: 'ref', node: { ...ref[0], sheet } })
  lex.at += ref[1]
}

function stringToken(lex: Lexer): boolean {
  if (lex.text[lex.at] !== '"') {
    return false
  }

  const end = quotedEnd(lex.text, lex.at, '"')
  lex.out.push({ kind: 'literal', value: lex.text.slice(lex.at + 1, end - 1).replace(/""/g, '"') })
  lex.at = end

  return true
}

function wordToken(lex: Lexer): boolean {
  const { text, at } = lex
  const word = matchAt(IDENTIFIER, text, at)
  if (!word) {
    return false
  }

  const end = at + word[0].length
  if (text[end] === '[') {
    throw new Unsupported('structured table references')
  }

  const upper = word[0].toUpperCase()
  if (text[end] === '(') {
    lex.out.push({ kind: 'func', text: upper.replace(/^(_XLFN\.|_XLWS\.)+/, '') })
  } else if (upper === 'TRUE' || upper === 'FALSE') {
    lex.out.push({ kind: 'literal', value: upper === 'TRUE' })
  } else {
    lex.out.push({ kind: 'name', text: word[0] })
  }

  lex.at = end

  return true
}

function valueToken(lex: Lexer): boolean {
  const { text, at } = lex
  const ref = readRef(text, at)
  if (ref) {
    lex.out.push({ kind: 'ref', node: ref[0] })
    lex.at += ref[1]
    return true
  }

  const number = matchAt(NUMBER, text, at)
  if (number) {
    lex.out.push({ kind: 'literal', value: Number(number[0]) })
    lex.at += number[0].length
    return true
  }

  const error = matchAt(ERROR_LITERAL, text, at)
  if (error) {
    lex.out.push({ kind: 'literal', value: errorNamed(error[0]) })
    lex.at += error[0].length
    return true
  }

  return false
}

function operatorToken(lex: Lexer): boolean {
  const op = [...OPERATORS, ...PUNCTUATION].find((o) => lex.text.startsWith(o, lex.at))
  if (!op) {
    return false
  }

  lex.out.push({ kind: 'op', text: op })
  lex.at += op.length

  return true
}

function step(lex: Lexer): void {
  if (/\s/.test(lex.text[lex.at])) {
    lex.at++
    return
  }

  if (lex.text[lex.at] === '[') {
    throw new Unsupported('external or structured references')
  }

  const sheet = sheetPrefix(lex)
  if (sheet !== null) {
    sheetToken(lex, sheet)
    return
  }

  if (stringToken(lex) || valueToken(lex) || wordToken(lex) || operatorToken(lex)) {
    return
  }

  throw new Unsupported(`unexpected ${lex.text[lex.at]}`)
}

export function tokenize(text: string): Token[] {
  const lex: Lexer = { text: text.replace(/^=/, ''), at: 0, out: [] }
  while (lex.at < lex.text.length) {
    step(lex)
  }

  return lex.out
}

class Parser {
  tokens: Token[]
  at: number

  constructor(tokens: Token[]) {
    this.tokens = tokens
    this.at = 0
  }

  peek(): Token | undefined {
    return this.tokens[this.at]
  }

  isOp(...ops: string[]): string | null {
    const token = this.peek()
    return token?.kind === 'op' && ops.includes(token.text) ? token.text : null
  }

  expect(op: string): void {
    if (!this.isOp(op)) {
      throw new Unsupported(`expected ${op}`)
    }

    this.at++
  }

  binary(ops: string[], next: () => Node): Node {
    let left = next()
    let op = this.isOp(...ops)
    while (op) {
      this.at++
      left = { type: 'binary', op, left, right: next() }
      op = this.isOp(...ops)
    }

    return left
  }

  comparison(): Node {
    return this.binary(['=', '<>', '<', '>', '<=', '>='], () => this.concat())
  }

  concat(): Node {
    return this.binary(['&'], () => this.additive())
  }

  additive(): Node {
    return this.binary(['+', '-'], () => this.multiplicative())
  }

  multiplicative(): Node {
    return this.binary(['*', '/'], () => this.power())
  }

  power(): Node {
    return this.binary(['^'], () => this.percent())
  }

  percent(): Node {
    let node = this.unary()
    while (this.isOp('%')) {
      this.at++
      node = { type: 'percent', arg: node }
    }

    return node
  }

  unary(): Node {
    const op = this.isOp('+', '-')
    if (op) {
      this.at++
      return { type: 'unary', op, arg: this.unary() }
    }

    return this.range()
  }

  range(): Node {
    let node = this.primary()
    while (this.isOp(':')) {
      this.at++
      node = { type: 'binary', op: ':', left: node, right: this.primary() }
    }

    return node
  }

  primary(): Node {
    const token = this.tokens[this.at++]
    if (!token) {
      throw new Unsupported('the formula ends early')
    }

    if (token.kind === 'literal') {
      return { type: 'literal', value: token.value }
    }

    if (token.kind === 'ref') {
      return token.node
    }

    if (token.kind === 'name') {
      return { type: 'name', name: token.text }
    }

    if (token.kind === 'func') {
      return this.call(token.text)
    }

    if (token.text === '(') {
      const inner = this.comparison()
      this.expect(')')
      return inner
    }

    if (token.text === '{') {
      return this.array()
    }

    throw new Unsupported(`unexpected ${token.text}`)
  }

  argument(): Node {
    return this.isOp(',', ')') ? { type: 'missing' } : this.comparison()
  }

  call(name: string): Node {
    this.expect('(')
    const args: Node[] = []
    if (this.isOp(')')) {
      this.at++
      return { type: 'call', name, args }
    }

    args.push(this.argument())
    while (this.isOp(',')) {
      this.at++
      args.push(this.argument())
    }

    this.expect(')')

    return { type: 'call', name, args }
  }

  constant(): Scalar {
    const token = this.tokens[this.at++]
    const sign = token?.kind === 'op' && token.text === '-' ? -1 : 1
    const value = sign === -1 ? this.tokens[this.at++] : token
    if (value?.kind !== 'literal' || (sign === -1 && typeof value.value !== 'number')) {
      throw new Unsupported('an array constant holds only values')
    }

    return typeof value.value === 'number' ? sign * value.value : value.value
  }

  array(): Node {
    const rows: Scalar[][] = [[this.constant()]]
    let sep = this.isOp(',', ';')
    while (sep) {
      this.at++
      if (sep === ';') {
        rows.push([])
      }

      rows[rows.length - 1].push(this.constant())
      sep = this.isOp(',', ';')
    }

    this.expect('}')
    if (rows.some((row) => row.length !== rows[0].length)) {
      throw new Unsupported('a ragged array constant')
    }

    return { type: 'array', rows }
  }
}

export function parse(text: string): Node {
  const parser = new Parser(tokenize(text))
  const node = parser.comparison()
  if (parser.at !== parser.tokens.length) {
    throw new Unsupported('the formula has more after its end')
  }

  return node
}
