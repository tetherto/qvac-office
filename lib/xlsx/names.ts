import * as xml from '../xml.ts'
import * as refs from './refs.ts'
import * as formula from './formula.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Element } from '../xml.ts'
import type { Area } from './refs.ts'
import type { ReshapeHost } from './reshape.ts'
import type { SheetEntry } from './sheet.ts'

export interface DefinedName {
  name: string
  ref: string
  sheet?: string
  hidden?: boolean
}

export interface NameOptions {
  sheet?: string | number
}

export interface NameDefined {
  name: string
  ref: string
  sheet?: string
  replaced: boolean
}

export interface NameDeleted {
  deleted: number
}

export interface NameRequest {
  name: unknown
  ref: unknown
  opts: unknown
}

interface Scope {
  index: number | null
  name?: string
}

interface Assignment {
  name: string
  ref: string
  scope: Scope
}

const X = xml.ns.x
const NAME = /^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u
const R1C1 = /^(?:[Rr]\d*(?:[Cc]\d*)?|[Cc]\d*)$/
const TABLE_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'
const FILTER_DATABASE = '_xlnm._FilterDatabase'
const AFTER_NAMES = [
  'calcPr',
  'oleSize',
  'customWorkbookViews',
  'pivotCaches',
  'smartTagPr',
  'smartTagTypes',
  'webPublishing',
  'fileRecoveryPr',
  'webPublishObjects',
  'extLst'
]

function looksLikeCell(name: string): boolean {
  try {
    refs.parse(name)
    return true
  } catch {
    return false
  }
}

export function validate(name: unknown, what = 'a defined name'): string {
  if (typeof name !== 'string' || name.length < 1 || name.length > 255) {
    throw new DocsError('INVALID_DEFINED_NAME', `${what} is 1 to 255 characters`, {
      field: 'name',
      got: name
    })
  }

  if (!NAME.test(name) || looksLikeCell(name) || R1C1.test(name)) {
    throw new DocsError(
      'INVALID_DEFINED_NAME',
      `"${name}" is not allowed as ${what}: it starts with a letter, _ or \\, holds only letters, digits, _ . and \\, and is not a cell reference like A1 or R1C1`,
      { field: 'name', got: name, expected: 'like Sales_2026 or TaxRate' }
    )
  }

  if (name.toLowerCase().startsWith('_xlnm.')) {
    throw new DocsError('INVALID_DEFINED_NAME', `"${name}" uses the _xlnm. prefix Excel reserves`)
  }

  return name
}

function container(host: ReshapeHost): Element {
  const root = host.package.root(host.main)
  const existing = xml.child(root, X, 'definedNames')
  if (existing) {
    return existing
  }

  const el = xml.named(host.package.xml(host.main), 'definedNames', X)
  const next = xml.children(root).find((n) => AFTER_NAMES.includes(n.localName ?? ''))
  root.insertBefore(el, next ?? null)

  return el
}

function scopeOf(dn: Element): number | null {
  const local = dn.getAttribute('localSheetId')
  return local === null || local === '' ? null : Number(local)
}

function scope(host: ReshapeHost, sheet: unknown): Scope {
  if (sheet === undefined) {
    return { index: null }
  }

  const entries = host.entries()
  const found =
    typeof sheet === 'number'
      ? entries[sheet]
      : entries.find(
          (e) => typeof sheet === 'string' && e.name.toLowerCase() === sheet.toLowerCase()
        )
  if (!found) {
    throw new DocsError(
      'SHEET_NOT_FOUND',
      `no sheet ${JSON.stringify(sheet)}; the workbook has ${entries.map((e) => `"${e.name}"`).join(', ')}`
    )
  }

  return { index: found.index, name: found.name }
}

export function list(host: ReshapeHost): DefinedName[] {
  const entries = host.entries()
  return host.definedNames().map((dn) => {
    const out: DefinedName = { name: xml.attr(dn, 'name'), ref: xml.text(dn) }
    const local = scopeOf(dn)
    if (local !== null) {
      out.sheet = entries[local]?.name ?? String(local)
    }

    if (dn.getAttribute('hidden') === '1' || dn.getAttribute('hidden') === 'true') {
      out.hidden = true
    }

    return out
  })
}

function find(host: ReshapeHost, name: string, index: number | null): Element | undefined {
  const wanted = name.toLowerCase()
  return host
    .definedNames()
    .find((dn) => xml.attr(dn, 'name').toLowerCase() === wanted && scopeOf(dn) === index)
}

function checkSheets(host: ReshapeHost, text: string): void {
  const names = new Set(host.entries().map((e) => e.name.toLowerCase()))
  const missing = formula.sheetsIn(text).filter((name) => !names.has(name.toLowerCase()))
  if (missing.length) {
    throw new DocsError(
      'SHEET_NOT_FOUND',
      `the reference names sheet ${missing.map((n) => `"${n}"`).join(', ')}, which does not exist; the workbook has ${host
        .entries()
        .map((e) => `"${e.name}"`)
        .join(', ')}`
    )
  }
}

function plainArea(text: string): Area | null {
  if (!/^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/.test(text)) {
    return null
  }

  return refs.parseRange(text.replace(/\$/g, ''))
}

function normalizeRef(host: ReshapeHost, ref: unknown, scoped: Scope): string {
  if (typeof ref !== 'string' || !ref.replace(/^=/, '').trim()) {
    throw new DocsError('INVALID_RANGE', 'ref is text like Sales!$A$1:$B$5 or =0.2')
  }

  const text = ref.replace(/^=/, '').trim()
  if (text.length > 8192) {
    throw new DocsError('INVALID_RANGE', 'ref is longer than 8192 characters')
  }

  const area = plainArea(text)
  if (area && !scoped.name) {
    throw new DocsError(
      'INVALID_RANGE',
      `${text} needs a sheet: write it like Sheet1!${text}, or give { sheet }`,
      { field: 'ref', got: text, expected: 'Sheet1!$A$1:$B$5' }
    )
  }

  if (area && scoped.name) {
    return refs.absolute(scoped.name, area)
  }

  checkSheets(host, text)

  return text
}

function assign(host: ReshapeHost, spec: Assignment): boolean {
  const existing = find(host, spec.name, spec.scope.index)
  const el = existing ?? xml.named(host.package.xml(host.main), 'definedName', X)
  el.setAttribute('name', existing ? xml.attr(existing, 'name') : spec.name)
  if (spec.scope.index !== null) {
    el.setAttribute('localSheetId', String(spec.scope.index))
  }

  xml.setText(el, spec.ref)
  if (!existing) {
    container(host).appendChild(el)
  }

  host.package.touch(host.main)

  return !!existing
}

export function tableNames(host: ReshapeHost): string[] {
  const pkg = host.package
  return pkg
    .names()
    .filter((part) => pkg.contentType(part) === TABLE_TYPE)
    .map((part) => pkg.root(part).getAttribute('displayName') || xml.attr(pkg.root(part), 'name'))
}

function checkTables(host: ReshapeHost, name: string): void {
  const clash = tableNames(host).find((t) => t.toLowerCase() === name.toLowerCase())
  if (clash) {
    throw new DocsError(
      'INVALID_DEFINED_NAME',
      `"${name}" is already the name of table ${clash}; tables and defined names share one namespace`
    )
  }
}

export function define(host: ReshapeHost, request: NameRequest): NameDefined {
  const { name, ref, opts } = request
  const valid = validate(name)
  checkTables(host, valid)
  const o = isRecord(opts) ? opts : {}
  const scoped = scope(host, o.sheet)
  const text = normalizeRef(host, ref, scoped)
  const replaced = assign(host, { name: valid, ref: text, scope: scoped })
  const out: NameDefined = { name: valid, ref: text, replaced }
  if (scoped.name) {
    out.sheet = scoped.name
  }

  return out
}

function dropEmpty(host: ReshapeHost): void {
  const names = xml.child(host.package.root(host.main), X, 'definedNames')
  if (names && !xml.children(names, X, 'definedName').length) {
    xml.remove(names)
  }
}

export function remove(host: ReshapeHost, name: unknown, opts: unknown): NameDeleted {
  const o = isRecord(opts) ? opts : {}
  const scoped = scope(host, o.sheet)
  const found = typeof name === 'string' ? find(host, name, scoped.index) : undefined
  if (!found) {
    const known = list(host).map((n) => (n.sheet ? `${n.sheet}!${n.name}` : n.name))
    throw new DocsError(
      'NAME_NOT_FOUND',
      `no defined name ${JSON.stringify(name)}${scoped.name ? ` on "${scoped.name}"` : ' at workbook scope'}; the workbook has ${known.length ? known.join(', ') : 'none'}`,
      { field: 'name', got: name, candidates: known }
    )
  }

  xml.remove(found)
  dropEmpty(host)
  host.package.touch(host.main)

  return { deleted: 1 }
}

export function setFilterDatabase(host: ReshapeHost, entry: SheetEntry, area: Area | null): void {
  const existing = find(host, FILTER_DATABASE, entry.index)
  if (!area) {
    if (existing) {
      xml.remove(existing)
      dropEmpty(host)
      host.package.touch(host.main)
    }

    return
  }

  assign(host, {
    name: FILTER_DATABASE,
    ref: refs.absolute(entry.name, area),
    scope: { index: entry.index }
  })
  find(host, FILTER_DATABASE, entry.index)?.setAttribute('hidden', '1')
}
