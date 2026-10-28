import * as xml from '../xml.ts'
import * as refs from './refs.ts'
import * as formula from './formula.ts'
import type { Document, Element } from '../xml.ts'
import type { Package } from '../opc/package.ts'
import type { Area } from './refs.ts'
import type { Change } from './formula.ts'
import type { EditState, Sheet, SheetEntry } from './sheet.ts'

export interface ReshapeHost {
  package: Package
  main: string
  state: EditState
  entries(): SheetEntry[]
  definedNames(): Element[]
  chartParts(): string[]
  drawingOf(sheet: Sheet): string | null
  sheetFor(entry: SheetEntry): Sheet
}

interface Table {
  part: string
  root: Element
  area: Area
  header: number
  totals: number
  name: string
}

const X = xml.ns.x
const XDR = xml.ns.xdr
const XM = 'http://schemas.microsoft.com/office/excel/2006/main'
const PIVOT_CACHE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml'
const COMMENTS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
const TABLE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
const PIVOT_TABLE_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable'

export function formulaElements(doc: Document): Element[] {
  return [
    ...xml.descendants(doc, X, 'f'),
    ...xml.descendants(doc, X, 'formula'),
    ...xml.descendants(doc, X, 'formula1'),
    ...xml.descendants(doc, X, 'formula2'),
    ...xml.descendants(doc, XM, 'f')
  ]
}

export function tableNames(pkg: Package, sheetPart: string): string[] {
  return pkg.related(sheetPart, TABLE_REL).map((part) => {
    const root = pkg.root(part)
    return root.getAttribute('displayName') || root.getAttribute('name') || part
  })
}

function tables(host: ReshapeHost, sheet: Sheet): Table[] {
  return host.package.related(sheet.part, TABLE_REL).map((part) => {
    const root = host.package.root(part)
    return {
      part,
      root,
      area: refs.parseRange(root.getAttribute('ref') || 'A1'),
      header: root.hasAttribute('headerRowCount') ? Number(root.getAttribute('headerRowCount')) : 1,
      totals: root.hasAttribute('totalsRowCount') ? Number(root.getAttribute('totalsRowCount')) : 0,
      name: root.getAttribute('displayName') || root.getAttribute('name') || part
    }
  })
}

function sqref(value: string, change: Change): string | null {
  const parts = value
    .trim()
    .split(/\s+/)
    .map((part) => formula.adjustRange(part, change))
    .filter((part) => part !== null)
  return parts.length ? parts.join(' ') : null
}

function tableBlockers(host: ReshapeHost, sheet: Sheet, change: Change): string[] {
  const reasons: string[] = []
  const end = change.at + change.count - 1
  const hits = (lo: number, hi: number): boolean => lo <= hi && change.at <= hi && end >= lo
  for (const table of tables(host, sheet)) {
    const { top, bottom } = table.area
    const dataTop = top + table.header
    const dataBottom = bottom - table.totals
    if (hits(top, dataTop - 1) || hits(dataBottom + 1, bottom)) {
      reasons.push(`the header or totals row of table ${table.name}`)
    } else if (change.at <= dataTop && end >= dataBottom && dataTop <= dataBottom) {
      reasons.push(`every data row of table ${table.name}`)
    }
  }

  return reasons
}

export function renamePivotSources(pkg: Package, from: string, to: string): void {
  const wanted = from.toLowerCase()
  for (const name of pkg.names().filter((n) => pkg.contentType(n) === PIVOT_CACHE)) {
    const sources = xml
      .descendants(pkg.xml(name), X, 'worksheetSource')
      .filter((src) => (src.getAttribute('sheet') || '').toLowerCase() === wanted)
    for (const src of sources) {
      src.setAttribute('sheet', to)
    }

    if (sources.length) {
      pkg.touch(name)
    }
  }
}

function hasPivot(host: ReshapeHost, sheet: Sheet): boolean {
  if (host.package.related(sheet.part, PIVOT_TABLE_REL).length) {
    return true
  }

  const target = sheet.name.toLowerCase()

  return host.package.names().some((name) => {
    if (host.package.contentType(name) !== PIVOT_CACHE) return false

    return xml
      .descendants(host.package.xml(name), X, 'worksheetSource')
      .some((src) => (src.getAttribute('sheet') || '').toLowerCase() === target)
  })
}

function straddles(area: Area, change: Change): boolean {
  const lo = change.axis === 'row' ? area.top : area.left
  const hi = change.axis === 'row' ? area.bottom : area.right
  if (change.kind === 'insert') {
    return change.at > lo && change.at <= hi
  }

  const past = change.at + change.count

  return change.at <= hi && past > lo && !(change.at <= lo && past > hi)
}

function formulaBlocker(sheet: Sheet, change: Change): string | null {
  for (const f of xml.descendants(sheet.doc, X, 'f')) {
    const kind = f.getAttribute('t')
    if (kind === 'dataTable') {
      return 'data tables'
    }

    const ref = f.getAttribute('ref')
    if (kind !== 'array' || !ref) {
      continue
    }

    if (straddles(refs.parseRange(ref), change)) {
      return `an array formula over ${ref}`
    }
  }

  return null
}

export function blockers(host: ReshapeHost, sheet: Sheet, change: Change): string[] {
  const reasons: string[] = []
  const root = sheet.root
  if (change.axis === 'col' && xml.child(root, X, 'tableParts')) {
    reasons.push('tables')
  }

  if (change.axis === 'row' && change.kind === 'delete') {
    reasons.push(...tableBlockers(host, sheet, change))
  }

  if (
    xml.child(root, X, 'legacyDrawing') ||
    host.package.related(sheet.part, COMMENTS_REL).length
  ) {
    reasons.push('comments')
  }

  if (hasPivot(host, sheet)) {
    reasons.push('pivot tables')
  }

  const blocker = formulaBlocker(sheet, change)
  if (blocker) {
    reasons.push(blocker)
  }

  return reasons
}

export function expandShared(sheet: Sheet): boolean {
  const shared = xml.descendants(sheet.doc, X, 'f').filter((f) => f.getAttribute('t') === 'shared')
  if (!shared.length) {
    return false
  }

  const texts = new Map<Element, string | null>()
  for (const f of shared) {
    const cell = f.parentNode
    texts.set(f, xml.isElement(cell) ? sheet.formula(cell) : null)
  }

  for (const f of shared) {
    const text = texts.get(f)
    for (const attr of ['t', 'ref', 'si']) {
      f.removeAttribute(attr)
    }

    if (text) {
      xml.setText(f, text.replace(/^=/, ''))
    }
  }

  sheet.resetMasters()
  sheet.touch()

  return true
}

function needsExpansion(sheet: Sheet, target: string): boolean {
  return xml
    .descendants(sheet.doc, X, 'f')
    .some(
      (f) =>
        f.getAttribute('t') === 'shared' &&
        !!xml.text(f) &&
        formula.references(xml.text(f), sheet.name, target)
    )
}

function adjustText(el: Element, ownSheet: string | null, change: Change): boolean {
  const text = xml.text(el)
  if (!text) {
    return false
  }

  const next = formula.adjust(text, ownSheet, change)
  if (next === text) {
    return false
  }

  xml.setText(el, next)

  return true
}

function adjustCellFormula(sheet: Sheet, f: Element, change: Change): boolean {
  let touched = adjustText(f, sheet.name, change)
  const ref = f.getAttribute('ref')
  if (sheet.name === change.sheet && ref) {
    const next = formula.adjustRange(ref, change)
    if (next && next !== ref) {
      f.setAttribute('ref', next)
      touched = true
    }
  }

  const cell = f.parentNode
  if (xml.isElement(cell) && formula.references(xml.text(f), sheet.name, change.sheet)) {
    for (const v of xml.children(cell, X, 'v')) {
      xml.remove(v)
    }

    cell.removeAttribute('t')
    touched = true
  }

  return touched
}

function adjustFormulas(sheet: Sheet, change: Change): boolean {
  let touched = false
  for (const f of xml.descendants(sheet.doc, X, 'f')) {
    if (adjustCellFormula(sheet, f, change)) {
      touched = true
    }
  }

  const rules = formulaElements(sheet.doc).filter(
    (el) => el.namespaceURI !== X || el.localName !== 'f'
  )
  for (const el of rules) {
    if (adjustText(el, sheet.name, change)) {
      touched = true
    }
  }

  return touched
}

function adjustList(list: Element | null, entry: string, change: Change): void {
  if (!list) return
  for (const el of xml.children(list, X, entry)) {
    const next = formula.adjustRange(xml.attr(el, 'ref'), change)
    const keep = entry === 'mergeCell' ? !!next && next.includes(':') : next !== null
    if (!keep || next === null) {
      xml.remove(el)
    } else {
      el.setAttribute('ref', next)
    }
  }

  const count = xml.children(list, X, entry).length
  if (!count) {
    xml.remove(list)
  } else if (entry === 'mergeCell') {
    list.setAttribute('count', String(count))
  }
}

function adjustSqrefs(sheet: Sheet, change: Change): void {
  for (const local of ['conditionalFormatting', 'dataValidation']) {
    for (const el of xml.descendants(sheet.doc, X, local)) {
      const value = el.getAttribute('sqref')
      if (value === null) {
        continue
      }

      const next = sqref(value, change)
      if (next === null) {
        xml.remove(el)
      } else {
        el.setAttribute('sqref', next)
      }
    }
  }

  const validations = xml.child(sheet.root, X, 'dataValidations')
  if (validations) {
    const count = xml.children(validations, X, 'dataValidation').length
    if (count) {
      validations.setAttribute('count', String(count))
    } else {
      xml.remove(validations)
    }
  }

  for (const el of xml.descendants(sheet.doc, XM, 'sqref')) {
    const next = sqref(xml.text(el), change)
    if (next !== null) {
      xml.setText(el, next)
    } else if (el.parentNode?.parentNode) {
      xml.remove(el.parentNode)
    }
  }
}

interface FilterMove {
  from: string
  to: string
  change: Change
}

function adjustFilterColumns(el: Element, move: FilterMove): void {
  const { change } = move
  if (change.axis !== 'col') {
    return
  }

  const left = refs.parseRange(move.from).left
  const nextLeft = refs.parseRange(move.to).left
  for (const column of xml.children(el, X, 'filterColumn')) {
    const moved = formula.moveIndex(left + Number(column.getAttribute('colId')), change)
    if (moved === null) {
      xml.remove(column)
    } else {
      column.setAttribute('colId', String(moved - nextLeft))
    }
  }
}

function adjustFilters(sheet: Sheet, change: Change): void {
  for (const local of ['autoFilter', 'sortState', 'sortCondition']) {
    for (const el of xml.descendants(sheet.doc, X, local)) {
      const ref = el.getAttribute('ref')
      if (ref === null) {
        continue
      }

      const next = formula.adjustRange(ref, change)
      if (next === null) {
        xml.remove(el)
      } else {
        adjustFilterColumns(el, { from: ref, to: next, change })
        el.setAttribute('ref', next)
      }
    }
  }
}

function adjustBreaks(sheet: Sheet, change: Change): void {
  const breaks = xml.child(sheet.root, X, change.axis === 'row' ? 'rowBreaks' : 'colBreaks')
  if (!breaks) {
    return
  }

  for (const brk of xml.children(breaks, X, 'brk')) {
    const next = formula.moveIndex(Number(brk.getAttribute('id')), change)
    if (next === null) {
      xml.remove(brk)
    } else {
      brk.setAttribute('id', String(next))
    }
  }

  const count = xml.children(breaks, X, 'brk').length
  if (!count) {
    xml.remove(breaks)
    return
  }

  breaks.setAttribute('count', String(count))
  if (breaks.hasAttribute('manualBreakCount')) {
    breaks.setAttribute('manualBreakCount', String(count))
  }
}

function adjustTables(host: ReshapeHost, sheet: Sheet, change: Change): void {
  if (change.axis !== 'row') return
  for (const table of tables(host, sheet)) {
    for (const el of [table.root, ...xml.children(table.root, X, 'autoFilter')]) {
      const next = formula.adjustRange(xml.attr(el, 'ref'), change)
      if (next) {
        el.setAttribute('ref', next)
      }
    }

    host.package.touch(table.part)
  }
}

function adjustAnchors(host: ReshapeHost, sheet: Sheet, change: Change): void {
  const drawing = host.drawingOf(sheet)
  if (!drawing) {
    return
  }

  const doc = host.package.xml(drawing)
  const local = change.axis === 'row' ? 'row' : 'col'
  let moved = false
  for (const marker of [...xml.descendants(doc, XDR, 'from'), ...xml.descendants(doc, XDR, 'to')]) {
    const el = xml.child(marker, XDR, local)
    if (!el) {
      continue
    }

    const index = Number(xml.text(el))
    const edge = (change.axis === 'row' ? refs.MAX_ROWS : refs.MAX_COLUMNS) - 1
    const value = formula.moveIndex(index, change) ?? (change.kind === 'insert' ? edge : change.at)
    if (value !== index) {
      xml.setText(el, String(value))
      moved = true
    }
  }

  if (moved) {
    host.package.touch(drawing)
  }
}

function adjustStructures(host: ReshapeHost, sheet: Sheet, change: Change): void {
  adjustList(xml.child(sheet.root, X, 'mergeCells'), 'mergeCell', change)
  adjustList(xml.child(sheet.root, X, 'hyperlinks'), 'hyperlink', change)
  adjustSqrefs(sheet, change)
  adjustFilters(sheet, change)
  adjustBreaks(sheet, change)
  adjustTables(host, sheet, change)
  adjustAnchors(host, sheet, change)
  sheet.touch()
}

function adjustElsewhere(host: ReshapeHost, change: Change): void {
  let names = false
  for (const dn of host.definedNames()) {
    if (adjustText(dn, null, change)) {
      names = true
    }
  }

  if (names) {
    host.package.touch(host.main)
  }

  for (const part of host.chartParts()) {
    let touched = false
    for (const f of xml.descendants(host.package.xml(part), xml.ns.c, 'f')) {
      if (adjustText(f, null, change)) {
        touched = true
      }
    }

    if (touched) {
      host.package.touch(part)
    }
  }
}

function sheetsOf(host: ReshapeHost, target: Sheet): Sheet[] {
  return host
    .entries()
    .filter((entry) => entry.kind === 'worksheet')
    .map((entry) => (entry.part === target.part ? target : host.sheetFor(entry)))
}

export function apply(host: ReshapeHost, target: Sheet, change: Change): void {
  for (const sheet of sheetsOf(host, target)) {
    if (adjustFormulas(sheet, change)) {
      sheet.touch()
    }
  }

  adjustStructures(host, target, change)
  adjustElsewhere(host, change)
  host.state.formulas = true
  host.state.cellsChanged = true
}

export function prepare(host: ReshapeHost, target: Sheet): void {
  for (const sheet of sheetsOf(host, target)) {
    if (sheet.part === target.part || needsExpansion(sheet, target.name)) {
      expandShared(sheet)
    }
  }
}
