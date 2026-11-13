import * as xml from '../xml.ts'
import * as refs from './refs.ts'
import * as areas from './areas.ts'
import * as dxf from './dxf.ts'
import { SHEET_ORDER, type Sheet } from './sheet.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Element } from '../xml.ts'
import type { Area } from './refs.ts'
import type { DiffStyle } from './dxf.ts'

export type CellIsOperator =
  | 'greaterThan'
  | 'lessThan'
  | 'greaterThanOrEqual'
  | 'lessThanOrEqual'
  | 'equal'
  | 'notEqual'
  | 'between'
  | 'notBetween'

export type RuleValue = number | string

export interface CellIsRule {
  kind: 'cellIs'
  operator: CellIsOperator
  value?: RuleValue
  values?: RuleValue[]
  style: DiffStyle
}

export interface ColorScaleRule {
  kind: 'colorScale'
  colors: string[]
}

export interface DataBarRule {
  kind: 'dataBar'
  color: string
}

export interface TextRule {
  kind: 'text'
  contains: string
  style: DiffStyle
}

export type ConditionalRule = CellIsRule | ColorScaleRule | DataBarRule | TextRule

export interface ConditionalAdded {
  range: string
  kind: string
  priority: number
}

export interface ConditionalDeleted {
  deleted: number
}

interface RuleContext {
  sheet: Sheet
  areas: Area[]
  priority: number
}

const X = xml.ns.x
const KINDS = ['cellIs', 'colorScale', 'dataBar', 'text']
const OPERATORS = [
  'greaterThan',
  'lessThan',
  'greaterThanOrEqual',
  'lessThanOrEqual',
  'equal',
  'notEqual',
  'between',
  'notBetween'
]

function formulaOf(value: unknown): string {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return String(value)
  }

  if (typeof value === 'string' && value.startsWith('=') && value.length > 1) {
    return value.slice(1)
  }

  if (typeof value === 'string') {
    return `"${value.replace(/"/g, '""')}"`
  }

  throw new DocsError(
    'INVALID_OPTIONS',
    `a rule value is a number, text, or a formula like "=$B$1", got ${String(value)}`
  )
}

function cellIsValues(rule: Record<string, unknown>, operator: string): string[] {
  const raw = rule.values !== undefined ? rule.values : [rule.value]
  const wanted = operator === 'between' || operator === 'notBetween' ? 2 : 1
  if (!Array.isArray(raw) || raw.length !== wanted || raw.some((v) => v === undefined)) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `${operator} takes ${wanted === 2 ? 'values: [low, high]' : 'one value'}`,
      { field: 'values', expected: wanted }
    )
  }

  return raw.map(formulaOf)
}

function cellIs(ctx: RuleContext, rule: Record<string, unknown>): Element {
  const operator = rule.operator
  if (typeof operator !== 'string' || !OPERATORS.includes(operator)) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `operator is one of ${OPERATORS.join(', ')}, got ${String(operator)}`,
      { field: 'operator', got: operator, candidates: OPERATORS }
    )
  }

  const values = cellIsValues(rule, operator)
  const d = xml.builder(ctx.sheet.doc, X)
  const dxfId = dxf.add(ctx.sheet.host.styles, rule.style)
  const el = d('cfRule', { type: 'cellIs', dxfId, priority: ctx.priority, operator })
  for (const value of values) {
    xml.setText(xml.append(el, d('formula')), value)
  }

  return el
}

function colorScale(ctx: RuleContext, rule: Record<string, unknown>): Element {
  const colors = rule.colors
  if (!Array.isArray(colors) || (colors.length !== 2 && colors.length !== 3)) {
    throw new DocsError('INVALID_OPTIONS', 'colors is 2 or 3 hex colors, low to high', {
      field: 'colors'
    })
  }

  const d = xml.builder(ctx.sheet.doc, X)
  const el = d('cfRule', { type: 'colorScale', priority: ctx.priority })
  const scale = xml.append(el, d('colorScale'))
  scale.appendChild(d('cfvo', { type: 'min' }))
  if (colors.length === 3) {
    scale.appendChild(d('cfvo', { type: 'percentile', val: 50 }))
  }

  scale.appendChild(d('cfvo', { type: 'max' }))
  colors.forEach((c: unknown, i: number) => {
    scale.appendChild(d('color', { rgb: dxf.color(c, `colors[${i}]`) }))
  })

  return el
}

function dataBar(ctx: RuleContext, rule: Record<string, unknown>): Element {
  const d = xml.builder(ctx.sheet.doc, X)
  const el = d('cfRule', { type: 'dataBar', priority: ctx.priority })
  const bar = xml.append(el, d('dataBar'))
  bar.appendChild(d('cfvo', { type: 'min' }))
  bar.appendChild(d('cfvo', { type: 'max' }))
  bar.appendChild(d('color', { rgb: dxf.color(rule.color, 'color') }))

  return el
}

function text(ctx: RuleContext, rule: Record<string, unknown>): Element {
  const contains = rule.contains
  if (typeof contains !== 'string' || !contains || contains.length > 255) {
    throw new DocsError('INVALID_OPTIONS', 'contains is the text to look for, 1 to 255 characters')
  }

  const d = xml.builder(ctx.sheet.doc, X)
  const dxfId = dxf.add(ctx.sheet.host.styles, rule.style)
  const first = refs.cell(ctx.areas[0].top, ctx.areas[0].left)
  const el = d('cfRule', {
    type: 'containsText',
    dxfId,
    priority: ctx.priority,
    operator: 'containsText',
    text: contains
  })
  const quoted = `"${contains.replace(/"/g, '""')}"`
  xml.setText(xml.append(el, d('formula')), `NOT(ISERROR(SEARCH(${quoted},${first})))`)

  return el
}

function nextPriority(sheet: Sheet): number {
  const all = xml
    .descendants(sheet.doc, X, 'cfRule')
    .map((r) => Number(r.getAttribute('priority')) || 0)
  return Math.max(0, ...all) + 1
}

function ruleElement(ctx: RuleContext, rule: unknown): Element {
  if (!isRecord(rule) || typeof rule.kind !== 'string' || !KINDS.includes(rule.kind)) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `rule is { kind } where kind is one of ${KINDS.join(', ')}`,
      { field: 'kind', candidates: KINDS }
    )
  }

  if (rule.kind === 'cellIs') {
    return cellIs(ctx, rule)
  }

  if (rule.kind === 'colorScale') {
    return colorScale(ctx, rule)
  }

  if (rule.kind === 'dataBar') {
    return dataBar(ctx, rule)
  }

  return text(ctx, rule)
}

export function add(sheet: Sheet, range: unknown, rule: unknown): ConditionalAdded {
  const list = areas.parseList(range)
  const ctx: RuleContext = { sheet, areas: list, priority: nextPriority(sheet) }
  const el = ruleElement(ctx, rule)
  const sqref = areas.format(list)
  const block = xml.setAttributes(xml.named(sheet.doc, 'conditionalFormatting', X), { sqref })
  block.appendChild(el)
  xml.insertOrdered(sheet.root, block, SHEET_ORDER)
  sheet.touch()

  return { range: sqref, kind: isRecord(rule) ? String(rule.kind) : '', priority: ctx.priority }
}

export function remove(sheet: Sheet, range: unknown): ConditionalDeleted {
  const target = range === undefined ? null : areas.parseList(range)
  let deleted = 0
  for (const block of xml.children(sheet.root, X, 'conditionalFormatting')) {
    const own = block.getAttribute('sqref')
    if (target && !(own && areas.anyIntersect(areas.parseList(own), target))) {
      continue
    }

    deleted += xml.children(block, X, 'cfRule').length
    xml.remove(block)
  }

  if (deleted) {
    sheet.touch()
  }

  return { deleted }
}
