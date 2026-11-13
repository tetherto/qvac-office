import * as xml from '../xml.ts'
import * as units from '../units.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import { validateColor, type ShapeEntry } from './shapes.ts'
import type { Element } from '../xml.ts'
import type { LineDash } from './layout-types.ts'

interface LineSpec {
  color?: string
  width?: number
  dash?: LineDash
}

const { a: A, p: P } = xml.ns
const FILLS = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']
const SP_PR_ORDER = [
  'xfrm',
  'custGeom',
  'prstGeom',
  ...FILLS,
  'ln',
  'effectLst',
  'effectDag',
  'scene3d',
  'sp3d',
  'extLst'
]
const LINE_ORDER = [
  'noFill',
  'solidFill',
  'gradFill',
  'pattFill',
  'prstDash',
  'custDash',
  'round',
  'bevel',
  'miter',
  'headEnd',
  'tailEnd',
  'extLst'
]
const DASHES: readonly LineDash[] = [
  'solid',
  'dot',
  'dash',
  'lgDash',
  'dashDot',
  'lgDashDot',
  'lgDashDotDot',
  'sysDash',
  'sysDot',
  'sysDashDot',
  'sysDashDotDot'
]
const MAX_LINE_WIDTH = 20116800
const FILLABLE = new Set(['sp', 'pic'])
const LINED = new Set(['sp', 'pic', 'cxnSp'])

function shapeProps(shape: ShapeEntry, allowed: Set<string>, what: string): Element {
  const local = shape.el.localName ?? ''
  if (!allowed.has(local)) {
    throw new DocsError(
      'NOT_SHAPE',
      `shape ${shape.id} is a ${shape.kind}, which takes no ${what}; use a shape, text box or picture`
    )
  }

  const existing = xml.child(shape.el, P, 'spPr')
  if (existing) {
    return existing
  }

  const made = xml.builder(xml.owner(shape.el), P)('spPr')
  shape.el.insertBefore(made, xml.children(shape.el)[1] ?? null)

  return made
}

function dropFills(parent: Element): void {
  for (const el of xml.children(parent, A)) {
    if (FILLS.includes(el.localName ?? '')) {
      xml.remove(el)
    }
  }
}

function fillElement(parent: Element, color: string | null): Element {
  const a = xml.builder(xml.owner(parent), A)
  if (color === null) {
    return a('noFill')
  }

  const fill = a('solidFill')
  fill.appendChild(a('srgbClr', { val: color.replace('#', '').toUpperCase() }))

  return fill
}

export function setFill(shape: ShapeEntry, color: unknown): void {
  const value = color === null ? null : validateColor(color)
  const spPr = shapeProps(shape, FILLABLE, 'fill')
  dropFills(spPr)
  xml.insertOrdered(spPr, fillElement(spPr, value), SP_PR_ORDER)
}

function lineWidth(value: unknown): number | undefined {
  const width = units.length(value, 'width')
  if (width === undefined) {
    return undefined
  }

  if (
    typeof width !== 'number' ||
    !Number.isInteger(width) ||
    width < 0 ||
    width > MAX_LINE_WIDTH
  ) {
    throw new DocsError(
      'INVALID_STYLE',
      `width must be a length from 0 to 1584pt, got ${String(value)}`,
      {
        field: 'width',
        got: value
      }
    )
  }

  return width
}

function lineSpec(input: unknown): LineSpec {
  if (!isRecord(input) || !Object.keys(input).length) {
    throw new DocsError('INVALID_STYLE', 'line must be { color, width, dash } or null')
  }

  const unknown = Object.keys(input).filter((key) => !['color', 'width', 'dash'].includes(key))
  if (unknown.length) {
    throw new DocsError(
      'INVALID_STYLE',
      `line takes color, width and dash, not ${unknown.join(', ')}`,
      {
        field: unknown[0],
        expected: ['color', 'width', 'dash']
      }
    )
  }

  const dash = input.dash
  if (dash !== undefined && !DASHES.some((d) => d === dash)) {
    throw new DocsError(
      'INVALID_STYLE',
      `dash must be one of ${DASHES.join(', ')}, got ${String(dash)}`,
      {
        field: 'dash',
        expected: DASHES,
        got: dash
      }
    )
  }

  return {
    color: input.color === undefined ? undefined : validateColor(input.color),
    width: lineWidth(input.width),
    dash: DASHES.find((d) => d === dash)
  }
}

function ensureLine(spPr: Element): Element {
  return (
    xml.child(spPr, A, 'ln') ??
    xml.insertOrdered(spPr, xml.builder(xml.owner(spPr), A)('ln'), SP_PR_ORDER)
  )
}

function writeLine(ln: Element, spec: LineSpec): void {
  if (spec.width !== undefined) {
    ln.setAttribute('w', String(spec.width))
  }

  if (spec.color !== undefined) {
    dropFills(ln)
    xml.insertOrdered(ln, fillElement(ln, spec.color), LINE_ORDER)
  }

  if (spec.dash !== undefined) {
    for (const el of [...xml.children(ln, A, 'prstDash'), ...xml.children(ln, A, 'custDash')]) {
      xml.remove(el)
    }

    xml.insertOrdered(ln, xml.builder(xml.owner(ln), A)('prstDash', { val: spec.dash }), LINE_ORDER)
  }
}

export function setLine(shape: ShapeEntry, input: unknown): void {
  const spec = input === null ? null : lineSpec(input)
  const spPr = shapeProps(shape, LINED, 'outline')
  if (spec) {
    writeLine(ensureLine(spPr), spec)
    return
  }

  const old = xml.child(spPr, A, 'ln')
  if (old) {
    xml.remove(old)
  }

  ensureLine(spPr).appendChild(fillElement(spPr, null))
}
