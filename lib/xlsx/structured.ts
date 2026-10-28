import * as xml from '../xml.ts'
import * as refs from './refs.ts'
import * as formula from './formula.ts'
import * as reshape from './reshape.ts'
import { DocsError } from '../errors.ts'
import type { Element } from '../xml.ts'
import type { Area } from './refs.ts'
import type { ReshapeHost } from './reshape.ts'
import type { TableEntry } from './tables.ts'

export interface Unlink {
  sheet: string
  table: TableEntry
  columns: string[]
}

interface Rows {
  top: number
  bottom: number
}

interface Site {
  el: Element
  part: string
  where: string
}

interface Rewrite {
  site: Site
  text: string | null
}

const SPECIAL = /^#(all|data|headers|totals|this row)$/i
const NAME_CHAR = /[\w.\\]/

export function escape(column: string): string {
  return column.replace(/(['#[\]])/g, "'$1")
}

function unescape(item: string): string {
  return item.trim().replace(/'(.)/g, '$1')
}

function rowsFor(special: string, spec: Unlink): Rows | null {
  const { area, header, totals } = spec.table
  const data = { top: area.top + header, bottom: area.bottom - totals }
  const key = special.toLowerCase()
  if (key === '#all') {
    return { top: area.top, bottom: area.bottom }
  }

  if (key === '#data') {
    return data
  }

  if (key === '#headers') {
    return header ? { top: area.top, bottom: area.top } : null
  }

  if (key === '#totals') {
    return totals ? { top: area.bottom, bottom: area.bottom } : null
  }

  return null
}

function items(inner: string): string[] {
  if (!inner.trim().startsWith('[')) {
    return [inner]
  }

  return Array.from(inner.matchAll(/\[((?:[^\]']|'.)*)\]/g), (m) => m[1])
}

function columnSpan(names: string[], spec: Unlink): [number, number] | null {
  const { left, right } = spec.table.area
  if (!names.length) {
    return [left, right]
  }

  const index = names.map((n) => spec.columns.findIndex((c) => c.toLowerCase() === n.toLowerCase()))
  if (index.some((i) => i === -1) || names.length > 2) {
    return null
  }

  return [left + Math.min(...index), left + Math.max(...index)]
}

function resolve(inner: string, spec: Unlink): Area | null {
  const parts = items(inner).map(unescape)
  if (parts.some((p) => p.startsWith('@'))) {
    return null
  }

  const specials = parts.filter((p) => SPECIAL.test(p))
  const columns = parts.filter((p) => !SPECIAL.test(p) && p !== '')
  const rows = (specials.length ? specials : ['#Data']).map((s) => rowsFor(s, spec))
  const span = columnSpan(columns, spec)
  if (rows.some((r) => r === null) || !span) {
    return null
  }

  const valid = rows.filter((r): r is Rows => r !== null)
  const top = Math.min(...valid.map((r) => r.top))
  const bottom = Math.max(...valid.map((r) => r.bottom))
  if (bottom < top) {
    return null
  }

  return { top, bottom, left: span[0], right: span[1] }
}

function bracketEnd(text: string, at: number): number {
  let depth = 0
  for (let i = at; i < text.length; i++) {
    if (text[i] === "'") {
      i++
    } else if (text[i] === '[') {
      depth++
    } else if (text[i] === ']' && --depth === 0) {
      return i + 1
    }
  }

  return -1
}

function startsTable(text: string, i: number, name: string): boolean {
  return (
    text.slice(i, i + name.length).toLowerCase() === name.toLowerCase() &&
    text[i + name.length] === '[' &&
    (i === 0 || !NAME_CHAR.test(text[i - 1]))
  )
}

export function convert(text: string, spec: Unlink): string | null {
  const name = spec.table.name
  let out = ''
  let i = 0
  while (i < text.length) {
    if (text[i] === '"') {
      const close = text.indexOf('"', i + 1)
      const end = close === -1 ? text.length : close + 1
      out += text.slice(i, end)
      i = end
    } else if (startsTable(text, i, name)) {
      const open = i + name.length
      const end = bracketEnd(text, open)
      const area = end === -1 ? null : resolve(text.slice(open + 1, end - 1), spec)
      if (!area) {
        return null
      }

      out += refs.absolute(spec.sheet, area)
      i = end
    } else {
      out += text[i]
      i++
    }
  }

  return out
}

function sites(host: ReshapeHost): Site[] {
  const out: Site[] = []
  for (const entry of host.entries().filter((e) => e.kind === 'worksheet')) {
    for (const el of reshape.formulaElements(host.package.xml(entry.part))) {
      const cell = el.parentNode
      const ref = xml.isElement(cell) ? cell.getAttribute('r') : null
      out.push({ el, part: entry.part, where: `${entry.name}!${ref ?? el.localName}` })
    }
  }

  for (const dn of host.definedNames()) {
    out.push({ el: dn, part: host.main, where: `defined name ${xml.attr(dn, 'name')}` })
  }

  return out
}

function plan(host: ReshapeHost, spec: Unlink): Rewrite[] {
  return sites(host)
    .filter((site) => formula.usesTable(xml.text(site.el), spec.table.name))
    .map((site) => ({ site, text: convert(xml.text(site.el), spec) }))
}

export function unlink(host: ReshapeHost, spec: Unlink): number {
  const rewrites = plan(host, spec)
  const stuck = rewrites.filter((r) => r.text === null).map((r) => r.site.where)
  if (stuck.length) {
    throw new DocsError(
      'TABLE_IS_REFERENCED',
      `table ${spec.table.name} is used by ${stuck.slice(0, 5).join(', ')}${stuck.length > 5 ? ` and ${stuck.length - 5} more` : ''} in a form that cannot be turned into a plain range; rewrite those formulas first`,
      { candidates: stuck }
    )
  }

  for (const { site, text } of rewrites) {
    xml.setText(site.el, text ?? '')
    host.package.touch(site.part)
  }

  if (rewrites.length) {
    host.state.formulas = true
  }

  return rewrites.length
}
