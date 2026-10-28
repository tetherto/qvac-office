import * as xml from '../xml.ts'
import * as areas from './areas.ts'
import { SHEET_ORDER, isoToSerial, type Sheet } from './sheet.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Element } from '../xml.ts'
import type { Area } from './refs.ts'

export type ValidationType = 'list' | 'whole' | 'decimal' | 'date' | 'textLength'

export type ValidationOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'lessThan'
  | 'greaterThanOrEqual'
  | 'lessThanOrEqual'

export interface ValidationMessage {
  title?: string
  text: string
}

export interface ValidationError extends ValidationMessage {
  style?: 'stop' | 'warning' | 'information'
}

export interface ValidationRule {
  type: ValidationType
  values?: string[] | string
  min?: number | string
  max?: number | string
  operator?: ValidationOperator
  allowBlank?: boolean
  prompt?: string | ValidationMessage
  error?: string | ValidationError
}

export interface ValidationAdded {
  range: string
  type: string
  operator?: string
}

export interface ValidationDeleted {
  deleted: number
}

interface Bounds {
  operator: string
  formulas: string[]
}

interface Context {
  sheet: Sheet
  type: string
}

const X = xml.ns.x
const TYPES = ['list', 'whole', 'decimal', 'date', 'textLength']
const OPERATORS = [
  'between',
  'notBetween',
  'equal',
  'notEqual',
  'greaterThan',
  'lessThan',
  'greaterThanOrEqual',
  'lessThanOrEqual'
]
const ERROR_STYLES = ['stop', 'warning', 'information']

function listFormula(values: unknown): string {
  if (typeof values === 'string' && values.startsWith('=') && values.length > 1) {
    return values.slice(1)
  }

  if (
    !Array.isArray(values) ||
    !values.length ||
    !values.every((v) => typeof v === 'string' && v && !v.includes(','))
  ) {
    throw new DocsError(
      'INVALID_OPTIONS',
      'a list takes values: ["Yes", "No"] (no commas inside an item) or a range like "=$H$1:$H$5"',
      { field: 'values' }
    )
  }

  const joined = values.join(',')
  if (joined.length > 255) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `the list is ${joined.length} characters; Excel takes at most 255, so put the items in cells and give their range`
    )
  }

  return `"${joined.replace(/"/g, '""')}"`
}

function bound(ctx: Context, value: unknown, field: string): string {
  if (typeof value === 'string' && value.startsWith('=') && value.length > 1) {
    return value.slice(1)
  }

  if (ctx.type === 'date' && typeof value === 'string') {
    return String(isoToSerial(value, ctx.sheet.host.date1904).serial)
  }

  const whole = ctx.type === 'whole' || ctx.type === 'textLength'
  const ok =
    typeof value === 'number' &&
    Number.isFinite(value) &&
    (!whole || Number.isInteger(value)) &&
    (ctx.type !== 'textLength' || value >= 0)
  if (!ok) {
    const expected =
      ctx.type === 'date' ? 'a date like "2026-01-31"' : `a ${whole ? 'whole ' : ''}number`
    throw new DocsError(
      'INVALID_OPTIONS',
      `${field} for ${ctx.type} is ${expected} or a formula like "=$B$1", got ${String(value)}`,
      { field, got: value, expected }
    )
  }

  return String(value)
}

function defaultOperator(o: Record<string, unknown>): string {
  if (o.min !== undefined && o.max !== undefined) return 'between'
  return o.min !== undefined ? 'greaterThanOrEqual' : 'lessThanOrEqual'
}

function bounds(ctx: Context, o: Record<string, unknown>): Bounds {
  if (o.min === undefined && o.max === undefined) {
    throw new DocsError('INVALID_OPTIONS', `${ctx.type} needs min, max or both`)
  }

  const operator = o.operator === undefined ? defaultOperator(o) : o.operator
  if (typeof operator !== 'string' || !OPERATORS.includes(operator)) {
    throw new DocsError('INVALID_OPTIONS', `operator is one of ${OPERATORS.join(', ')}`, {
      field: 'operator',
      got: operator,
      candidates: OPERATORS
    })
  }

  if (operator === 'between' || operator === 'notBetween') {
    if (o.min === undefined || o.max === undefined) {
      throw new DocsError('INVALID_OPTIONS', `${operator} needs both min and max`)
    }

    return { operator, formulas: [bound(ctx, o.min, 'min'), bound(ctx, o.max, 'max')] }
  }

  const single = o.min === undefined ? 'max' : 'min'

  return { operator, formulas: [bound(ctx, o[single], single)] }
}

function message(value: unknown, field: string): ValidationError | null {
  if (value === undefined) return null
  const msg = typeof value === 'string' ? { text: value } : value
  if (!isRecord(msg) || typeof msg.text !== 'string' || msg.text.length > 255) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `${field} is text up to 255 characters, or { title, text }`
    )
  }

  if (msg.title !== undefined && (typeof msg.title !== 'string' || msg.title.length > 32)) {
    throw new DocsError('INVALID_OPTIONS', `${field}.title is text up to 32 characters`)
  }

  if (msg.style !== undefined && !ERROR_STYLES.includes(String(msg.style))) {
    throw new DocsError('INVALID_OPTIONS', `error.style is one of ${ERROR_STYLES.join(', ')}`)
  }

  return {
    text: msg.text,
    title: typeof msg.title === 'string' ? msg.title : undefined,
    style: msg.style === 'warning' || msg.style === 'information' ? msg.style : undefined
  }
}

function guardOverlap(sheet: Sheet, target: Area[]): void {
  for (const el of xml.descendants(sheet.doc, X, 'dataValidation')) {
    const own = el.getAttribute('sqref')
    if (own && areas.anyIntersect(areas.parseList(own), target)) {
      throw new DocsError(
        'VALIDATION_OVERLAP',
        `${areas.format(target)} overlaps the validation on ${own}; delete it first with deleteValidations`,
        { field: 'range', candidates: [own] }
      )
    }
  }
}

function container(sheet: Sheet): Element {
  return (
    xml.child(sheet.root, X, 'dataValidations') ??
    xml.insertOrdered(sheet.root, xml.named(sheet.doc, 'dataValidations', X), SHEET_ORDER)
  )
}

function element(sheet: Sheet, o: Record<string, unknown>, sqref: string): Element {
  const type = String(o.type)
  const ctx = { sheet, type }
  const spec =
    type === 'list' ? { operator: '', formulas: [listFormula(o.values)] } : bounds(ctx, o)
  const prompt = message(o.prompt, 'prompt')
  const error = message(o.error, 'error')
  const d = xml.builder(sheet.doc, X)
  const el = d('dataValidation', {
    type,
    errorStyle: error?.style,
    operator: spec.operator || undefined,
    allowBlank: o.allowBlank === false ? undefined : 1,
    showInputMessage: prompt ? 1 : undefined,
    showErrorMessage: 1,
    errorTitle: error?.title,
    error: error?.text,
    promptTitle: prompt?.title,
    prompt: prompt?.text,
    sqref
  })
  spec.formulas.forEach((f, i) => {
    xml.setText(xml.append(el, d(`formula${i + 1}`)), f)
  })

  return el
}

export function add(sheet: Sheet, range: unknown, rule: unknown): ValidationAdded {
  if (!isRecord(rule) || typeof rule.type !== 'string' || !TYPES.includes(rule.type)) {
    throw new DocsError('INVALID_OPTIONS', `type is one of ${TYPES.join(', ')}`, {
      field: 'type',
      candidates: TYPES
    })
  }

  const target = areas.parseList(range)
  guardOverlap(sheet, target)
  const sqref = areas.format(target)
  const el = element(sheet, rule, sqref)
  const list = container(sheet)
  list.appendChild(el)
  list.setAttribute('count', String(xml.children(list, X, 'dataValidation').length))
  sheet.touch()
  const out: ValidationAdded = { range: sqref, type: rule.type }
  if (el.getAttribute('operator')) {
    out.operator = xml.attr(el, 'operator')
  }

  return out
}

export function remove(sheet: Sheet, range: unknown): ValidationDeleted {
  const target = range === undefined ? null : areas.parseList(range)
  const list = xml.child(sheet.root, X, 'dataValidations')
  if (!list) {
    return { deleted: 0 }
  }

  let deleted = 0
  for (const el of xml.children(list, X, 'dataValidation')) {
    const own = el.getAttribute('sqref')
    if (target && !(own && areas.anyIntersect(areas.parseList(own), target))) {
      continue
    }

    xml.remove(el)
    deleted++
  }

  if (!deleted) {
    return { deleted }
  }

  const left = xml.children(list, X, 'dataValidation').length
  if (left) {
    list.setAttribute('count', String(left))
  } else {
    xml.remove(list)
  }

  sheet.touch()

  return { deleted }
}
