import * as xml from '../xml.ts'
import * as text from './text.ts'
import * as units from '../units.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import { PAGES } from './template.ts'
import { W, SECTION_ORDER, body, touchMain, type Host } from './host.ts'
import type { Element } from '../xml.ts'

export type Orientation = 'portrait' | 'landscape'
export type PageName = 'letter' | 'a4'

export interface PageSize {
  width: number
  height: number
}

export interface PageMargins {
  top: number
  right: number
  bottom: number
  left: number
  header: number
  footer: number
}

export interface PageSetup {
  size: PageSize
  name: PageName | null
  orientation: Orientation
  margins: PageMargins
  sections: number
}

export interface PageOptions {
  size?: PageName | { width: units.Length; height: units.Length }
  orientation?: Orientation
  margins?: Partial<Record<keyof PageMargins, units.Length>>
  all?: boolean
}

export interface PageResult {
  sections: number
  page: PageSetup
}

interface Twips {
  width: number
  height: number
}

const EMU_PER_TWIP = 635
const MARGIN_KEYS = ['top', 'right', 'bottom', 'left', 'header', 'footer'] as const
const MARGIN_DEFAULTS: Record<string, number> = {
  top: 1440,
  right: 1440,
  bottom: 1440,
  left: 1440,
  header: 720,
  footer: 720,
  gutter: 0
}
const MARGIN_NAMES: ReadonlySet<string> = new Set(MARGIN_KEYS)
const PAGE_NAMES: readonly PageName[] = ['letter', 'a4']
const ORIENTATIONS: readonly string[] = ['portrait', 'landscape']

function sections(host: Host): Element[] {
  return xml
    .descendants(body(host), W, 'sectPr')
    .filter((el) => !xml.isElement(el.parentNode, W, 'sectPrChange'))
}

function finalSection(host: Host): Element {
  const all = sections(host)
  const last = all[all.length - 1]
  if (!last) {
    throw new DocsError('BROKEN_DOCX', 'the document has no section properties')
  }

  return last
}

function twips(emu: unknown, field: string, min: number): number {
  if (!(typeof emu === 'number' && Number.isInteger(emu) && emu >= min)) {
    throw new DocsError(
      'INVALID_BOX',
      `${field} must be a whole number of EMU or a length such as "1in", got ${String(emu)}`,
      { field, expected: min > 0 ? 'positive length' : 'length of 0 or more', got: emu }
    )
  }

  return Math.round(emu / EMU_PER_TWIP)
}

function namedSize(name: unknown): Twips {
  const page = typeof name === 'string' ? PAGES[name.toLowerCase()] : undefined
  if (!page) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `size must be "letter", "a4" or { width, height }, got ${String(name)}`,
      {
        field: 'size',
        expected: ['letter', 'a4', '{ width, height }'],
        got: name
      }
    )
  }

  return { ...page }
}

function sizeOf(input: unknown): Twips | null {
  if (input === undefined) {
    return null
  }

  if (typeof input === 'string' || !isRecord(input)) {
    return namedSize(input)
  }

  return {
    width: twips(units.length(input.width, 'width'), 'width', 1),
    height: twips(units.length(input.height, 'height'), 'height', 1)
  }
}

function checkOrientation(value: unknown): Orientation | null {
  if (value === undefined) {
    return null
  }

  if (value === 'portrait' || value === 'landscape') {
    return value
  }

  throw new DocsError(
    'INVALID_OPTIONS',
    `orientation must be "portrait" or "landscape", got ${String(value)}`,
    {
      field: 'orientation',
      expected: ORIENTATIONS,
      got: value
    }
  )
}

function marginsOf(input: unknown): Record<string, number> {
  if (input === undefined) {
    return {}
  }

  if (!isRecord(input)) {
    throw new DocsError(
      'INVALID_OPTIONS',
      'margins must be { top, right, bottom, left, header, footer }'
    )
  }

  const unknownKey = Object.keys(input).find((key) => !MARGIN_NAMES.has(key))
  if (unknownKey) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `unknown margin "${unknownKey}"; use ${MARGIN_KEYS.join(', ')}`,
      {
        field: 'margins',
        expected: MARGIN_KEYS,
        got: unknownKey
      }
    )
  }

  return Object.fromEntries(
    Object.entries(input).map(([key, value]) => [key, twips(units.length(value, key), key, 0)])
  )
}

function ensure(sectPr: Element, local: string): Element {
  return (
    xml.child(sectPr, W, local) ??
    xml.insertOrdered(sectPr, text.wEl(xml.owner(sectPr), local), SECTION_ORDER)
  )
}

function current(pgSz: Element): Twips {
  return {
    width: Number(text.wGet(pgSz, 'w')) || PAGES.letter.width,
    height: Number(text.wGet(pgSz, 'h')) || PAGES.letter.height
  }
}

function oriented(size: Twips, orientation: Orientation): Twips {
  const wide = size.width > size.height
  const swap = orientation === 'landscape' ? !wide : wide
  return swap ? { width: size.height, height: size.width } : size
}

interface Change {
  size: Twips | null
  named: boolean
  orientation: Orientation | null
  margins: Record<string, number>
}

function applySize(sectPr: Element, change: Change): void {
  if (!change.size && !change.orientation) {
    return
  }

  const pgSz = ensure(sectPr, 'pgSz')
  const base = change.size ?? current(pgSz)
  const orientation =
    change.orientation ?? (text.wGet(pgSz, 'orient') === 'landscape' ? 'landscape' : 'portrait')
  const size =
    change.size && !change.named && !change.orientation ? base : oriented(base, orientation)
  text.wAttr(pgSz, 'w', size.width)
  text.wAttr(pgSz, 'h', size.height)
  if (size.width > size.height) {
    text.wAttr(pgSz, 'orient', 'landscape')
  } else {
    pgSz.removeAttributeNS(W, 'orient')
  }
}

function applyMargins(sectPr: Element, margins: Record<string, number>): void {
  if (!Object.keys(margins).length) {
    return
  }

  const pgMar = ensure(sectPr, 'pgMar')
  for (const [key, fallback] of Object.entries(MARGIN_DEFAULTS)) {
    const value = margins[key] ?? (text.wGet(pgMar, key) || fallback)
    text.wAttr(pgMar, key, value)
  }
}

export function set(host: Host, opts: PageOptions): PageResult {
  if (!isRecord(opts)) {
    throw new DocsError('INVALID_OPTIONS', 'give the page as { size, orientation, margins, all }')
  }

  const change: Change = {
    size: sizeOf(opts.size),
    named: typeof opts.size === 'string',
    orientation: checkOrientation(opts.orientation),
    margins: marginsOf(opts.margins)
  }
  const targets = opts.all ? sections(host) : [finalSection(host)]
  for (const sectPr of targets) {
    applySize(sectPr, change)
    applyMargins(sectPr, change.margins)
  }

  touchMain(host)

  return { sections: targets.length, page: read(host) }
}

function nameOf(size: Twips): PageName | null {
  for (const name of PAGE_NAMES) {
    const same = oriented(PAGES[name], size.width > size.height ? 'landscape' : 'portrait')
    if (Math.abs(same.width - size.width) <= 2 && Math.abs(same.height - size.height) <= 2) {
      return name
    }
  }

  return null
}

export function read(host: Host): PageSetup {
  const sectPr = finalSection(host)
  const pgSz = xml.child(sectPr, W, 'pgSz')
  const pgMar = xml.child(sectPr, W, 'pgMar')
  const size = pgSz ? current(pgSz) : { ...PAGES.letter }
  const margin = (key: string): number =>
    (pgMar && text.wGet(pgMar, key) ? Number(text.wGet(pgMar, key)) : MARGIN_DEFAULTS[key]) *
    EMU_PER_TWIP
  const landscape = pgSz
    ? text.wGet(pgSz, 'orient') === 'landscape' || size.width > size.height
    : false

  return {
    size: { width: size.width * EMU_PER_TWIP, height: size.height * EMU_PER_TWIP },
    name: nameOf(size),
    orientation: landscape ? 'landscape' : 'portrait',
    margins: {
      top: margin('top'),
      right: margin('right'),
      bottom: margin('bottom'),
      left: margin('left'),
      header: margin('header'),
      footer: margin('footer')
    },
    sections: sections(host).length
  }
}
