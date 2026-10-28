import * as xml from '../xml.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Element } from '../xml.ts'
import type { Styles } from './styles.ts'

export interface DiffStyle {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  color?: string
  fill?: string
}

const X = xml.ns.x
const KEYS = ['bold', 'italic', 'underline', 'color', 'fill']
const STYLE_ORDER = [
  'numFmts',
  'fonts',
  'fills',
  'borders',
  'cellStyleXfs',
  'cellXfs',
  'cellStyles',
  'dxfs',
  'tableStyles',
  'colors',
  'extLst'
]
const HEX = /^#?[0-9A-Fa-f]{6}$/

export function color(value: unknown, field: string): string {
  if (typeof value !== 'string' || !HEX.test(value)) {
    throw new DocsError(
      'INVALID_STYLE',
      `${field} is a hex color like "C0392B", got ${String(value)}`,
      { field, got: value, expected: 'RRGGBB' }
    )
  }

  return 'FF' + value.replace('#', '').toUpperCase()
}

function validate(style: unknown): Record<string, unknown> {
  if (!isRecord(style) || Array.isArray(style) || !Object.keys(style).length) {
    throw new DocsError(
      'INVALID_STYLE',
      'style names what the rule changes, like { fill: "FFC7CE", color: "9C0006" }'
    )
  }

  const unknown = Object.keys(style).filter((k) => !KEYS.includes(k))
  if (unknown.length) {
    throw new DocsError(
      'INVALID_STYLE',
      `a rule style takes ${KEYS.join(', ')}; unknown ${unknown.join(', ')}`,
      { field: 'style', got: unknown, candidates: KEYS }
    )
  }

  for (const key of ['bold', 'italic', 'underline']) {
    if (style[key] !== undefined && typeof style[key] !== 'boolean') {
      throw new DocsError('INVALID_STYLE', `${key} is true or false`)
    }
  }

  return style
}

function build(styles: Styles, style: Record<string, unknown>): Element {
  const d = xml.builder(styles.doc, X)
  const dxf = d('dxf')
  if (style.bold || style.italic || style.underline || style.color !== undefined) {
    const font = xml.append(dxf, d('font'))
    if (style.bold !== undefined) {
      font.appendChild(d('b', { val: style.bold ? 1 : 0 }))
    }

    if (style.italic !== undefined) {
      font.appendChild(d('i', { val: style.italic ? 1 : 0 }))
    }

    if (style.underline) {
      font.appendChild(d('u'))
    }

    if (style.color !== undefined) {
      font.appendChild(d('color', { rgb: color(style.color, 'color') }))
    }
  }

  if (style.fill !== undefined) {
    const pattern = xml.append(xml.append(dxf, d('fill')), d('patternFill'))
    pattern.appendChild(d('bgColor', { rgb: color(style.fill, 'fill') }))
  }

  return dxf
}

function list(styles: Styles): Element {
  const top = xml.root(styles.doc)
  return (
    xml.child(top, X, 'dxfs') ??
    xml.insertOrdered(
      top,
      xml.setAttributes(xml.named(styles.doc, 'dxfs', X), { count: 0 }),
      STYLE_ORDER
    )
  )
}

export function add(styles: Styles, input: unknown): number {
  const dxf = build(styles, validate(input))
  const dxfs = list(styles)
  const siblings = xml.children(dxfs, X, 'dxf')
  const key = dxf.toString()
  const existing = siblings.findIndex((el) => el.toString() === key)
  if (existing !== -1) {
    return existing
  }

  dxfs.appendChild(dxf)
  dxfs.setAttribute('count', String(siblings.length + 1))
  styles.package.touch(styles.part ?? '')

  return siblings.length
}
