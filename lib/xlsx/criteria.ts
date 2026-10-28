import { Unsupported, compare, isError, rounded, textToNumber } from './values.ts'
import type { Scalar } from './values.ts'

export interface Criterion {
  test(value: Scalar): boolean
  blank: boolean
}

const OPERATOR = /^(<=|>=|<>|=|<|>)?([\s\S]*)$/
const ORDERINGS: Record<string, (order: number) => boolean> = {
  '<': (c) => c < 0,
  '>': (c) => c > 0,
  '<=': (c) => c <= 0,
  '>=': (c) => c >= 0
}
const ERROR_TEXT = /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A)$/i

export function looseNumber(value: Scalar): number | null {
  if (typeof value === 'number') return value
  if (typeof value !== 'string') return null
  try {
    return textToNumber(value)
  } catch (err) {
    if (err instanceof Unsupported) {
      return null
    }

    throw err
  }
}

export function wildcard(pattern: string): RegExp {
  let source = ''
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i]
    if (ch === '~' && i + 1 < pattern.length && '*?~'.includes(pattern[i + 1])) {
      source += '\\' + pattern[++i]
    } else if (ch === '*') {
      source += '[\\s\\S]*'
    } else if (ch === '?') {
      source += '[\\s\\S]'
    } else {
      source += ch.replace(/[.+^${}()|[\]\\/-]/g, '\\$&')
    }
  }

  return new RegExp(`^${source}$`, 'i')
}

function equality(rest: string): (value: Scalar) => boolean {
  const n = looseNumber(rest)
  if (n !== null) {
    return (value) => {
      const v = looseNumber(value)
      return v !== null && rounded(v) === rounded(n)
    }
  }

  const upper = rest.toUpperCase()
  if (upper === 'TRUE' || upper === 'FALSE') {
    return (value) => value === (upper === 'TRUE')
  }

  if (ERROR_TEXT.test(rest)) {
    return (value) => isError(value) && value.code === upper
  }

  const re = wildcard(rest)

  return (value) => typeof value === 'string' && re.test(value)
}

function ordering(op: string, rest: string): (value: Scalar) => boolean {
  const n = looseNumber(rest)
  const upper = rest.toUpperCase()
  const logical = upper === 'TRUE' || upper === 'FALSE'
  const fallback: Scalar = logical ? upper === 'TRUE' : rest
  const target: Scalar = n ?? fallback
  const passes = ORDERINGS[op]

  return (value) =>
    value !== null && typeof value === typeof target && passes(compare(value, target))
}

function fromText(text: string): Criterion {
  const [, op = '', rest] = OPERATOR.exec(text) ?? ['', '', text]
  if (op === '' || op === '=') {
    if (rest === '') {
      return { test: (value) => op === '' && value === '', blank: true }
    }

    return { test: equality(rest), blank: false }
  }

  if (op === '<>') {
    if (rest === '') {
      return { test: () => true, blank: false }
    }

    const equal = equality(rest)

    return { test: (value) => !equal(value), blank: true }
  }

  return { test: ordering(op, rest), blank: false }
}

export function criterion(value: Scalar): Criterion {
  if (typeof value === 'string') {
    return fromText(value)
  }

  if (typeof value === 'number') {
    return fromText(String(value))
  }

  if (typeof value === 'boolean') {
    return { test: (v) => v === value, blank: false }
  }

  if (isError(value)) {
    return { test: (v) => isError(v) && v.code === value.code, blank: false }
  }

  throw new Unsupported('a blank criterion')
}
