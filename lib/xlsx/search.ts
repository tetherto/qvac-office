import * as refs from './refs.ts'
import { DocsError } from '../errors.ts'
import { isRecord, errorMessage } from '../guards.ts'
import type { CellValue, Sheet } from './sheet.ts'
import type { ReshapeHost } from './reshape.ts'

export interface FindOptions {
  sheet?: string | number
  regex?: boolean
  matchCase?: boolean
  formulas?: boolean
  limit?: number
}

export interface Found {
  sheet: string
  ref: string
  value: CellValue
  formula?: string
}

export interface SearchScope {
  host: ReshapeHost
  sheet: (ref: string | number) => Sheet
}

type Matcher = (text: string) => boolean

function matcher(query: unknown, o: Record<string, unknown>): Matcher {
  if (typeof query !== 'string' || !query) {
    throw new DocsError('INVALID_OPTIONS', 'find takes the text to look for')
  }

  if (o.regex === true) {
    let re: RegExp
    try {
      re = new RegExp(query, o.matchCase === true ? 'u' : 'iu')
    } catch (err) {
      throw new DocsError('INVALID_OPTIONS', `${query} is not a valid regular expression`, {
        field: 'query',
        got: query,
        cause: errorMessage(err)
      })
    }

    return (text) => re.test(text)
  }

  if (o.matchCase === true) {
    return (text) => text.includes(query)
  }

  const lower = query.toLowerCase()

  return (text) => text.toLowerCase().includes(lower)
}

function display(value: CellValue): string {
  if (value === null) return ''
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE'
  }

  return String(value)
}

function limitOf(value: unknown): number {
  if (value === undefined) return 1000
  if (!(typeof value === 'number' && Number.isInteger(value) && value > 0)) {
    throw new DocsError('INVALID_OPTIONS', `limit is a positive whole number, got ${String(value)}`)
  }

  return value
}

function sheetsIn(scope: SearchScope, ref: unknown): Sheet[] {
  if (typeof ref === 'string' || typeof ref === 'number') {
    return [scope.sheet(ref)]
  }

  return scope.host
    .entries()
    .filter((e) => e.kind === 'worksheet')
    .map((e) => scope.host.sheetFor(e))
}

export function find(scope: SearchScope, query: unknown, opts: unknown): Found[] {
  const o = isRecord(opts) ? opts : {}
  const matches = matcher(query, o)
  const limit = limitOf(o.limit)
  const out: Found[] = []
  for (const sheet of sheetsIn(scope, o.sheet)) {
    for (const { el, row, col } of sheet.cells()) {
      const value = sheet.value(el)
      const formula = sheet.formula(el)
      const hit = matches(display(value)) || (o.formulas === true && !!formula && matches(formula))
      if (!hit) {
        continue
      }

      const found: Found = { sheet: sheet.name, ref: refs.cell(row, col), value }
      if (formula) {
        found.formula = formula
      }

      out.push(found)
      if (out.length >= limit) {
        return out
      }
    }
  }

  return out
}
