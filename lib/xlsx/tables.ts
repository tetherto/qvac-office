import * as xml from '../xml.ts'
import * as refs from './refs.ts'
import * as areas from './areas.ts'
import * as names from './names.ts'
import * as structured from './structured.ts'
import { SHEET_ORDER, type Sheet, type CellValue } from './sheet.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Element } from '../xml.ts'
import type { Area } from './refs.ts'
import type { ReshapeHost } from './reshape.ts'

export interface TableEntry {
  part: string
  root: Element
  name: string
  ref: string
  area: Area
  header: number
  totals: number
}

export interface TableInfo {
  name: string
  sheet: string
  range: string
  header: boolean
  totals: boolean
  columns: string[]
  style?: string
}

export interface TableOptions {
  name?: string
  header?: boolean
  style?: string
  totals?: boolean
}

export interface AddedTable {
  name: string
  range: string
  columns: string[]
}

export interface DeletedTable {
  deleted: string
  converted: number
}

export interface TableRequest {
  range: unknown
  opts: unknown
}

interface TableSpec {
  name: string
  area: Area
  header: boolean
  totals: boolean
  style: string
  columns: string[]
}

interface Located {
  sheet: Sheet
  table: TableEntry
}

const X = xml.ns.x
const R = xml.ns.r
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
const TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'

function entryOf(sheet: Sheet, part: string): TableEntry {
  const root = sheet.package.root(part)
  const ref = root.getAttribute('ref') || 'A1'
  return {
    part,
    root,
    name: root.getAttribute('displayName') || root.getAttribute('name') || part,
    ref,
    area: refs.parseRange(ref),
    header: root.hasAttribute('headerRowCount') ? Number(root.getAttribute('headerRowCount')) : 1,
    totals: root.hasAttribute('totalsRowCount') ? Number(root.getAttribute('totalsRowCount')) : 0
  }
}

export function on(sheet: Sheet): TableEntry[] {
  return sheet.package.related(sheet.part, REL).map((part) => entryOf(sheet, part))
}

function worksheets(host: ReshapeHost): Sheet[] {
  return host
    .entries()
    .filter((e) => e.kind === 'worksheet')
    .map((e) => host.sheetFor(e))
}

function info(sheet: Sheet, table: TableEntry): TableInfo {
  const columns = xml.descendants(table.root, X, 'tableColumn').map((c) => xml.attr(c, 'name'))
  const out: TableInfo = {
    name: table.name,
    sheet: sheet.name,
    range: table.ref,
    header: table.header > 0,
    totals: table.totals > 0,
    columns
  }
  const style = xml.child(table.root, X, 'tableStyleInfo')?.getAttribute('name')
  if (style) {
    out.style = style
  }

  return out
}

export function list(host: ReshapeHost, only: Sheet | null): TableInfo[] {
  const sheets = only ? [only] : worksheets(host)
  return sheets.flatMap((sheet) => on(sheet).map((table) => info(sheet, table)))
}

function headerText(value: CellValue, i: number): string {
  if (value === null || value === '') {
    return `Column${i + 1}`
  }

  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE'
  }

  return String(value)
    .replace(/[\r\n]+/g, ' ')
    .slice(0, 255)
}

function columnNames(sheet: Sheet, spec: TableSpec): string[] {
  const seen = new Set<string>()
  const out: string[] = []
  for (let col = spec.area.left; col <= spec.area.right; col++) {
    const i = col - spec.area.left
    const cell = spec.header ? sheet.findCell({ row: spec.area.top, col }) : null
    const base = headerText(cell ? sheet.value(cell) : null, i)
    let name = base
    for (let n = 2; seen.has(name.toLowerCase()); n++) {
      name = `${base}${n}`
    }

    seen.add(name.toLowerCase())
    out.push(name)
  }

  return out
}

function freeName(host: ReshapeHost): string {
  const taken = takenNames(host)
  let n = 1
  while (taken.has(`table${n}`)) {
    n++
  }

  return `Table${n}`
}

function takenNames(host: ReshapeHost): Set<string> {
  return new Set(
    [...names.tableNames(host), ...host.definedNames().map((dn) => xml.attr(dn, 'name'))].map((n) =>
      n.toLowerCase()
    )
  )
}

function chooseName(host: ReshapeHost, name: unknown): string {
  if (name === undefined) {
    return freeName(host)
  }

  const valid = names.validate(name, 'a table name')
  if (takenNames(host).has(valid.toLowerCase())) {
    throw new DocsError(
      'INVALID_DEFINED_NAME',
      `"${valid}" is already a table or defined name; pick another or leave name out for ${freeName(host)}`,
      { field: 'name', got: valid }
    )
  }

  return valid
}

function bool(value: unknown, field: string): boolean | undefined {
  if (value === undefined) {
    return undefined
  }

  if (typeof value !== 'boolean') {
    throw new DocsError('INVALID_OPTIONS', `${field} is true or false, got ${String(value)}`)
  }

  return value
}

function styleName(value: unknown): string {
  if (value === undefined) return 'TableStyleMedium2'
  if (typeof value !== 'string' || !/^[A-Za-z][\w ]{0,254}$/.test(value)) {
    throw new DocsError(
      'INVALID_STYLE',
      `style is a table style name like "TableStyleMedium2" or "TableStyleLight9", got ${String(value)}`
    )
  }

  return value
}

function guardArea(sheet: Sheet, spec: TableSpec): void {
  const ref = refs.range(spec.area)
  const minRows = (spec.header ? 2 : 1) + (spec.totals ? 1 : 0)
  if (spec.area.bottom - spec.area.top + 1 < minRows || spec.area.bottom >= refs.MAX_ROWS) {
    throw new DocsError(
      'INVALID_RANGE',
      `${ref} is too short: a table needs ${spec.header ? 'a header row and ' : ''}at least one data row`
    )
  }

  const merged = sheet.merged().filter((m) => areas.intersects(refs.parseRange(m), spec.area))
  if (merged.length) {
    throw new DocsError(
      'RANGE_HAS_MERGES',
      `${ref} overlaps the merged range ${merged.join(', ')}; unmerge it first`,
      { candidates: merged }
    )
  }

  const other = on(sheet).find((t) => areas.intersects(t.area, spec.area))
  if (other) {
    throw new DocsError('TABLE_OVERLAP', `${ref} overlaps table ${other.name} at ${other.ref}`)
  }

  const filter = xml.child(sheet.root, X, 'autoFilter')?.getAttribute('ref')
  if (filter && areas.intersects(refs.parseRange(filter), spec.area)) {
    throw new DocsError(
      'TABLE_OVERLAP',
      `${ref} overlaps the sheet's autofilter on ${filter}; remove it with autofilter(sheet, null) first`
    )
  }
}

function guardTotalsRow(sheet: Sheet, spec: TableSpec): void {
  if (!spec.totals) return
  const row = spec.area.bottom
  const used = sheet
    .cells()
    .find(
      (c) =>
        c.row === row &&
        c.col >= spec.area.left &&
        c.col <= spec.area.right &&
        (!!c.el.firstChild || c.el.hasAttribute('t'))
    )
  if (used) {
    throw new DocsError(
      'INVALID_RANGE',
      `the totals row goes below the range, in row ${row + 1}, but ${refs.cell(row, used.col)} is not empty; insert a row first`
    )
  }
}

function writeHeaders(sheet: Sheet, spec: TableSpec): void {
  if (!spec.header) return
  spec.columns.forEach((name, i) => {
    const pos = { row: spec.area.top, col: spec.area.left + i }
    const cell = sheet.findCell(pos)
    if (!cell || sheet.value(cell) !== name || sheet.formula(cell)) {
      sheet.write(refs.cell(pos.row, pos.col), { value: name })
    }
  })
}

function writeTotals(sheet: Sheet, spec: TableSpec): void {
  if (!spec.totals) return
  const row = spec.area.bottom
  sheet.write(refs.cell(row, spec.area.left), { value: 'Total' })
  if (spec.columns.length > 1) {
    const last = spec.columns[spec.columns.length - 1]
    sheet.write(refs.cell(row, spec.area.right), {
      formula: `SUBTOTAL(109,${spec.name}[${structured.escape(last)}])`
    })
  }
}

function nextId(host: ReshapeHost): number {
  const pkg = host.package
  const ids = pkg
    .names()
    .filter((part) => pkg.contentType(part) === TYPE)
    .map((part) => Number(pkg.root(part).getAttribute('id')) || 0)

  return Math.max(0, ...ids) + 1
}

function tableXml(host: ReshapeHost, spec: TableSpec): xml.Document {
  const doc = xml.parse(`<table xmlns="${X}"/>`)
  const t = xml.builder(doc, X)
  const root = xml.setAttributes(xml.root(doc), {
    id: nextId(host),
    name: spec.name,
    displayName: spec.name,
    ref: refs.range(spec.area),
    headerRowCount: spec.header ? undefined : 0,
    totalsRowCount: spec.totals ? 1 : undefined
  })
  if (spec.header) {
    const bottom = spec.area.bottom - (spec.totals ? 1 : 0)
    root.appendChild(t('autoFilter', { ref: refs.range({ ...spec.area, bottom }) }))
  }

  const cols = xml.append(root, t('tableColumns', { count: spec.columns.length }))
  spec.columns.forEach((name, i) => {
    const last = i === spec.columns.length - 1 && i > 0
    cols.appendChild(
      t('tableColumn', {
        id: i + 1,
        name,
        totalsRowLabel: spec.totals && i === 0 ? 'Total' : undefined,
        totalsRowFunction: spec.totals && last ? 'sum' : undefined
      })
    )
  })
  root.appendChild(
    t('tableStyleInfo', {
      name: spec.style,
      showFirstColumn: 0,
      showLastColumn: 0,
      showRowStripes: 1,
      showColumnStripes: 0
    })
  )

  return doc
}

function attach(host: ReshapeHost, sheet: Sheet, spec: TableSpec): void {
  const part = host.package.freeName('/xl/tables/table%d.xml')
  host.package.add(part, tableXml(host, spec), TYPE)
  const rId = host.package.rels(sheet.part).add(REL, part)
  const list =
    xml.child(sheet.root, X, 'tableParts') ??
    xml.insertOrdered(sheet.root, xml.named(sheet.doc, 'tableParts', X), SHEET_ORDER)
  const el = xml.named(sheet.doc, 'tablePart', X)
  xml.setNsAttribute(el, { namespace: R, local: 'id' }, rId)
  list.appendChild(el)
  list.setAttribute('count', String(xml.children(list, X, 'tablePart').length))
  sheet.touch()
}

function specOf(host: ReshapeHost, range: unknown, opts: unknown): TableSpec {
  if (opts !== undefined && (!isRecord(opts) || Array.isArray(opts))) {
    throw new DocsError('INVALID_OPTIONS', 'addTable takes { name?, header?, style?, totals? }')
  }

  const o = isRecord(opts) ? opts : {}
  const totals = bool(o.totals, 'totals') ?? false
  const given = areas.parse(range)

  return {
    name: chooseName(host, o.name),
    area: totals ? { ...given, bottom: given.bottom + 1 } : given,
    header: bool(o.header, 'header') ?? true,
    totals,
    style: styleName(o.style),
    columns: []
  }
}

export function add(host: ReshapeHost, sheet: Sheet, request: TableRequest): AddedTable {
  const spec = specOf(host, request.range, request.opts)
  guardArea(sheet, spec)
  guardTotalsRow(sheet, spec)
  spec.columns = columnNames(sheet, spec)
  writeHeaders(sheet, spec)
  writeTotals(sheet, spec)
  attach(host, sheet, spec)

  return { name: spec.name, range: refs.range(spec.area), columns: spec.columns }
}

function locate(host: ReshapeHost, name: unknown): Located {
  const wanted = typeof name === 'string' ? name.toLowerCase() : null
  for (const sheet of worksheets(host)) {
    const table = on(sheet).find((t) => t.name.toLowerCase() === wanted)
    if (table) {
      return { sheet, table }
    }
  }

  const known = list(host, null).map((t) => t.name)

  throw new DocsError(
    'TABLE_NOT_FOUND',
    `no table ${JSON.stringify(name)}; the workbook has ${known.length ? known.join(', ') : 'no tables'}`,
    { field: 'name', got: name, candidates: known }
  )
}

function detach(host: ReshapeHost, found: Located): void {
  const { sheet, table } = found
  const rels = host.package.rels(sheet.part)
  const rel = rels.all().find((r) => r.type === REL && r.part === table.part)
  const list = xml.child(sheet.root, X, 'tableParts')
  for (const el of list ? xml.children(list, X, 'tablePart') : []) {
    if (rel && el.getAttributeNS(R, 'id') === rel.id) {
      xml.remove(el)
    }
  }

  if (list && !xml.children(list, X, 'tablePart').length) {
    xml.remove(list)
  } else if (list) {
    list.setAttribute('count', String(xml.children(list, X, 'tablePart').length))
  }

  if (rel) {
    rels.remove(rel.id)
  }

  host.package.remove(table.part)
  sheet.touch()
}

export function remove(host: ReshapeHost, name: unknown): DeletedTable {
  const found = locate(host, name)
  const columns = xml
    .descendants(found.table.root, X, 'tableColumn')
    .map((c) => xml.attr(c, 'name'))
  const converted = structured.unlink(host, {
    sheet: found.sheet.name,
    table: found.table,
    columns
  })
  detach(host, found)

  return { deleted: found.table.name, converted }
}
