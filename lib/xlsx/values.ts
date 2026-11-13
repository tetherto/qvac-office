import { isRecord } from '../guards.ts'
import type { Area } from './refs.ts'

export class ErrorValue {
  code: string

  constructor(code: string) {
    this.code = code
  }
}

export class Unsupported extends Error {}

export type Scalar = number | string | boolean | ErrorValue | null

export interface RangeValue {
  kind: 'range'
  sheet: string
  area: Area
}

export interface ArrayValue {
  kind: 'array'
  rows: Scalar[][]
}

export type Value = Scalar | RangeValue | ArrayValue

export const ERRORS = {
  null: new ErrorValue('#NULL!'),
  div0: new ErrorValue('#DIV/0!'),
  value: new ErrorValue('#VALUE!'),
  ref: new ErrorValue('#REF!'),
  name: new ErrorValue('#NAME?'),
  num: new ErrorValue('#NUM!'),
  na: new ErrorValue('#N/A')
}

const KNOWN_ERRORS = new Map(Object.values(ERRORS).map((e) => [e.code, e]))
const NUMERIC_TEXT = /^[+-]?(?=\.?\d)(\d{1,3}(,\d{3})+|\d+)?(\.\d*)?([eE][+-]?\d+)?%?$/
const MAYBE_VALUE = /[$€£¥()/:]|\d\s*-\s*\d|\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/i
const BIG = 1e15
const TINY = 1e-9

export function errorNamed(code: string): ErrorValue {
  const known = KNOWN_ERRORS.get(code.toUpperCase())
  if (!known) {
    throw new Unsupported(code)
  }

  return known
}

export function isError(value: unknown): value is ErrorValue {
  return value instanceof ErrorValue
}

export function isRange(value: Value): value is RangeValue {
  return isRecord(value) && Reflect.get(value, 'kind') === 'range'
}

export function isArray(value: Value): value is ArrayValue {
  return isRecord(value) && Reflect.get(value, 'kind') === 'array'
}

export function rounded(n: number): number {
  return n === 0 || !Number.isFinite(n) ? n : Number(n.toPrecision(15))
}

export function textToNumber(text: string): number | null {
  const trimmed = text.trim()
  if (!/\d/.test(trimmed)) {
    return null
  }

  if (!NUMERIC_TEXT.test(trimmed)) {
    if (MAYBE_VALUE.test(trimmed)) {
      throw new Unsupported(`text "${text}" may be a date, time or currency`)
    }

    return null
  }

  const percent = trimmed.endsWith('%')
  const n = Number(trimmed.replace(/[,%]/g, ''))
  if (!Number.isFinite(n)) {
    return null
  }

  return percent ? n / 100 : n
}

export function toNumber(value: Scalar): number | ErrorValue {
  if (typeof value === 'number') return value
  if (value === null) return 0
  if (typeof value === 'boolean') {
    return value ? 1 : 0
  }

  if (isError(value)) {
    return value
  }

  return textToNumber(value) ?? ERRORS.value
}

function exponentText(n: number): string {
  const [mantissa, exponent] = n.toExponential(14).split('e')
  const digits = mantissa.replace(/\.?0+$/, '')
  const power = Number(exponent)
  const sign = power < 0 ? '-' : '+'

  return `${digits}E${sign}${String(Math.abs(power)).padStart(2, '0')}`
}

export function numberText(n: number): string {
  const abs = Math.abs(n)
  if (abs === 0) {
    return '0'
  }

  if (abs >= BIG) {
    return exponentText(n)
  }

  if (abs < TINY) {
    throw new Unsupported('a very small number has no fixed text form')
  }

  return String(rounded(n))
}

export function toText(value: Scalar): string | ErrorValue {
  if (typeof value === 'string') return value
  if (value === null) return ''
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE'
  }

  if (isError(value)) {
    return value
  }

  return numberText(value)
}

export function toBool(value: Scalar): boolean | ErrorValue {
  if (typeof value === 'boolean') return value
  if (value === null) return false
  if (typeof value === 'number') {
    return value !== 0
  }

  if (isError(value)) {
    return value
  }

  const upper = value.trim().toUpperCase()
  if (upper === 'TRUE') {
    return true
  }

  if (upper === 'FALSE') {
    return false
  }

  return ERRORS.value
}

function rank(value: Scalar): number {
  if (typeof value === 'number') return 0
  if (typeof value === 'string') return 1
  return 2
}

function blankAs(other: Scalar): Scalar {
  if (typeof other === 'string') return ''
  if (typeof other === 'boolean') return false
  return 0
}

function order<T extends number | string>(x: T, y: T): number {
  if (x === y) {
    return 0
  }

  return x < y ? -1 : 1
}

export function compare(a: Scalar, b: Scalar): number {
  const left = a === null ? blankAs(b) : a
  const right = b === null ? blankAs(a) : b
  if (rank(left) !== rank(right)) {
    return rank(left) - rank(right)
  }

  if (typeof left === 'number' && typeof right === 'number') {
    return order(rounded(left), rounded(right))
  }

  if (typeof left === 'string' && typeof right === 'string') {
    return order(left.toLowerCase(), right.toLowerCase())
  }

  return Number(left) - Number(right)
}

export function checked(n: number): number | ErrorValue {
  return Number.isFinite(n) ? n : ERRORS.num
}
