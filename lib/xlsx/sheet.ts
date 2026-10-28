import * as xml from '../xml.ts'
import * as refs from './refs.ts'
import * as formulas from './formula.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Document, Element } from '../xml.ts'
import type { Package } from '../opc/package.ts'
import { runText, type SharedStrings } from './strings.ts'
import type { Styles } from './styles.ts'
import type { Area, Position } from './refs.ts'

export type CellValue = string | number | boolean | null

export interface EditState {
  touched: Set<Sheet>
  formulas: boolean
  cellsChanged: boolean
}

export interface SheetHost {
  package: Package
  strings: SharedStrings
  styles: Styles
  state: EditState
  readonly date1904: boolean
}

export interface SheetEntry {
  el: Element
  index: number
  rId: string
  name: string
  state: string
  part: string
  kind: string
}

export interface CellEntry extends Position {
  el: Element
}

export interface RowEntry {
  el: Element
  r: number
}

export interface ReadOptions {
  range?: string
  limit?: number
}

export interface SheetRead {
  name: string
  range: string | null
  rows: CellValue[][]
  formulas: Record<string, string>
  merged: string[]
  truncated?: boolean
}

export interface CellRead {
  value: CellValue
  formula: string | null
  format: string
}

export interface SerialDate {
  serial: number
  time: boolean
}

export interface ColumnWidth {
  col: number
  width: number
}

interface CellSpec {
  content: boolean
  value?: CellValue
  formula?: string
  date?: string
  style: Record<string, unknown>
}

interface Removal {
  at: number
  count: number
}

interface Master {
  text: string
  row: number
  col: number
}

const X = xml.ns.x
const DAY = 86400000
const EPOCH_1900 = Date.UTC(1899, 11, 30)
const EPOCH_1900_EARLY = Date.UTC(1899, 11, 31)
const EPOCH_1904 = Date.UTC(1904, 0, 1)

export const SHEET_ORDER = [
  'sheetPr',
  'dimension',
  'sheetViews',
  'sheetFormatPr',
  'cols',
  'sheetData',
  'sheetCalcPr',
  'sheetProtection',
  'protectedRanges',
  'scenarios',
  'autoFilter',
  'sortState',
  'dataConsolidate',
  'customSheetViews',
  'mergeCells',
  'phoneticPr',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'drawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst'
]

function timeOf(serial: number): string {
  return new Date(Math.round((serial % 1) * DAY)).toISOString().slice(11, 19)
}

export function serialToIso(serial: number, date1904: boolean): string | number {
  if (!(serial >= 0 && serial <= 2958465.99999)) return serial
  if (!date1904 && Math.floor(serial) === 60) {
    return Number.isInteger(serial) ? '1900-02-29' : `1900-02-29T${timeOf(serial)}`
  }

  let epoch = date1904 ? EPOCH_1904 : EPOCH_1900
  if (!date1904 && serial < 61) {
    epoch = EPOCH_1900_EARLY
  }

  const iso = new Date(Math.round(serial * DAY) + epoch).toISOString()

  return Number.isInteger(serial) ? iso.slice(0, 10) : iso.slice(0, 19)
}

function serialFrom(ms: number, date1904: boolean): number {
  if (date1904) {
    return (ms - EPOCH_1904) / DAY
  }

  if (ms < Date.UTC(1900, 2, 1)) {
    return (ms - EPOCH_1900_EARLY) / DAY
  }

  return (ms - EPOCH_1900) / DAY
}

export function isoToSerial(value: string, date1904: boolean): SerialDate {
  const match = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?$/.exec(value)
  if (!match) {
    throw new DocsError(
      'INVALID_VALUE',
      `date must be YYYY-MM-DD or YYYY-MM-DDTHH:MM, got ${value}`
    )
  }

  const [, y, mo, d, h = '0', mi = '0', s = '0'] = match
  const ms = Date.UTC(Number(y), Number(mo) - 1, Number(d), Number(h), Number(mi), Number(s))
  const check = new Date(ms)
  if (check.getUTCMonth() !== Number(mo) - 1 || check.getUTCDate() !== Number(d)) {
    throw new DocsError('INVALID_VALUE', `${value} is not a real date`)
  }

  const serial = serialFrom(ms, date1904)
  if (serial < (date1904 ? 0 : 1)) {
    throw new DocsError('INVALID_VALUE', `${value} is before the workbook's first date`)
  }

  return { serial, time: match[4] !== undefined }
}

function formulaError(input: string, why: string): DocsError {
  return new DocsError(
    'INVALID_FORMULA',
    `${JSON.stringify(input)} is not a formula (${why}); to store it as text use { value: ${JSON.stringify(input)} }`
  )
}

function parenthesisProblem(body: string): string | null {
  let depth = 0
  let i = 0
  while (i < body.length) {
    const ch = body[i]
    if (ch === '"' || ch === "'") {
      const close = body.indexOf(ch, i + 1)
      if (close === -1) {
        return `an unclosed ${ch === '"' ? 'string' : 'sheet name'}`
      }

      i = close + 1
      continue
    }

    if (ch === '(') {
      depth++
    } else if (ch === ')' && --depth < 0) {
      return 'a ) without a matching ('
    }

    i++
  }

  return depth > 0 ? 'a ( without a matching )' : null
}

function checkFormula(input: string): void {
  const body = input.replace(/^=/, '')
  if (!body.trim()) {
    throw formulaError(input, 'it is empty')
  }

  if (body.length > 8192) {
    throw formulaError(input, 'it is longer than 8192 characters')
  }

  const problem = parenthesisProblem(body)
  if (problem) {
    throw formulaError(input, problem)
  }

  if (/[+\-*/^&=<>,(]\s*$/.test(body)) {
    throw formulaError(input, 'it ends with an operator')
  }
}

function isCellValue(value: unknown): value is CellValue {
  return (
    value === null ||
    typeof value === 'string' ||
    typeof value === 'boolean' ||
    (typeof value === 'number' && Number.isFinite(value))
  )
}

function normalizeText(input: string): CellSpec {
  if (input.length > 32767) {
    throw new DocsError('INVALID_VALUE', 'a cell holds at most 32767 characters')
  }

  if (input.length > 1 && input.startsWith('=')) {
    checkFormula(input)
    return { content: true, formula: input, style: {} }
  }

  return { content: true, value: input, style: {} }
}

function normalizeObject(input: Record<string, unknown>): CellSpec {
  const { value, formula, date, ...style } = input
  const given = [value, formula, date].filter((v) => v !== undefined).length
  if (given > 1) {
    throw new DocsError('INVALID_VALUE', 'give only one of value, formula or date')
  }

  if (formula !== undefined) {
    if (typeof formula !== 'string') {
      throw new DocsError('INVALID_VALUE', 'formula must be text like "=SUM(A1:A3)"')
    }

    checkFormula(formula)

    return { content: true, formula, style }
  }

  if (date !== undefined) {
    if (typeof date !== 'string') {
      throw new DocsError('INVALID_VALUE', 'date must be text like "2026-03-31"')
    }

    return { content: true, date, style }
  }

  if (value === undefined) {
    return { content: false, style }
  }

  if (!isCellValue(value)) {
    throw new DocsError(
      'INVALID_VALUE',
      'value must be a string, a finite number, a boolean or null'
    )
  }

  return { content: true, value, style }
}

function normalize(input: unknown): CellSpec {
  if (input === null || input === undefined) {
    return { content: true, value: null, style: {} }
  }

  if (typeof input === 'number') {
    if (!Number.isFinite(input)) {
      throw new DocsError('INVALID_VALUE', 'numbers must be finite')
    }

    return { content: true, value: input, style: {} }
  }

  if (typeof input === 'boolean') {
    return { content: true, value: input, style: {} }
  }

  if (typeof input === 'string') {
    return normalizeText(input)
  }

  if (isRecord(input) && !Array.isArray(input)) {
    return normalizeObject(input)
  }

  throw new DocsError(
    'INVALID_VALUE',
    'a cell takes a string, number, boolean, null or { value | formula | date, ...style }'
  )
}

interface Shrink {
  at: number
  end: number
  count: number
}

function shrunkEnd(max: number, shrink: Shrink): number {
  if (max < shrink.at) return max
  return max <= shrink.end ? shrink.at - 1 : max - shrink.count
}

function rowNumber(row: Element): number {
  return Number(row.getAttribute('r'))
}

function within(pos: Position, area: Area): boolean {
  return (
    pos.row >= area.top && pos.row <= area.bottom && pos.col >= area.left && pos.col <= area.right
  )
}

function parentCell(f: Element): Element {
  const parent = f.parentNode
  if (!xml.isElement(parent)) {
    throw new DocsError('BROKEN_XLSX', 'a formula sits outside a cell')
  }

  return parent
}

function columnSpan(el: Element, n: number): void {
  const cols = el.parentNode
  const min = Number(el.getAttribute('min'))
  const max = Number(el.getAttribute('max'))
  if (!cols || n < min || n > max) {
    return
  }

  if (min < n) {
    const before = xml.clone(el)
    before.setAttribute('max', String(n - 1))
    cols.insertBefore(before, el)
  }

  if (max > n) {
    const after = xml.clone(el)
    after.setAttribute('min', String(n + 1))
    cols.insertBefore(after, el.nextSibling)
  }

  xml.remove(el)
}

export class Sheet {
  host: SheetHost
  name: string
  part: string
  index: number
  private masters: Map<string, Master> | null

  constructor(host: SheetHost, entry: SheetEntry) {
    this.host = host
    this.name = entry.name
    this.part = entry.part
    this.index = entry.index
    this.masters = null
  }

  get package(): Package {
    return this.host.package
  }

  get doc(): Document {
    return this.package.xml(this.part)
  }

  get root(): Element {
    return xml.root(this.doc)
  }

  sheetData(): Element {
    const existing = xml.child(this.root, X, 'sheetData')
    if (existing) {
      return existing
    }

    return xml.insertOrdered(this.root, xml.named(this.doc, 'sheetData', X), SHEET_ORDER)
  }

  rows(): Element[] {
    return xml.children(this.sheetData(), X, 'row')
  }

  touch(): void {
    this.package.touch(this.part)
    this.host.state.touched.add(this)
  }

  cells(): CellEntry[] {
    return this.numberedRows().flatMap(({ el, r }) => this.numberedCells(el, r))
  }

  numberedRows(): RowEntry[] {
    let next = 0
    return this.rows().map((el) => {
      const r = el.hasAttribute('r') ? rowNumber(el) - 1 : next
      next = r + 1

      return { el, r }
    })
  }

  numberedCells(row: Element, r: number): CellEntry[] {
    let lastCol = -1
    return xml.children(row, X, 'c').map((c) => {
      const pos = c.hasAttribute('r')
        ? this.position(xml.attr(c, 'r'))
        : { row: r, col: lastCol + 1 }
      lastCol = pos.col

      return { el: c, row: pos.row, col: pos.col }
    })
  }

  number(): void {
    for (const { el, row, col } of this.cells()) {
      if (!el.hasAttribute('r')) {
        el.setAttribute('r', refs.cell(row, col))
      }
    }

    for (const { el, r } of this.numberedRows()) {
      if (!el.hasAttribute('r')) {
        el.setAttribute('r', String(r + 1))
      }
    }
  }

  positionOf(c: Element): Position {
    if (c.hasAttribute('r')) {
      return this.position(xml.attr(c, 'r'))
    }

    const found = this.cells().find((entry) => entry.el === c)
    if (!found) {
      throw new DocsError('BROKEN_XLSX', `sheet "${this.name}" has a cell outside any row`)
    }

    return { row: found.row, col: found.col }
  }

  usedRange(): Area | null {
    let area: Area | null = null
    for (const { el, row, col } of this.cells()) {
      if (!el.firstChild && !el.hasAttribute('t')) continue
      area = area
        ? {
            top: Math.min(area.top, row),
            left: Math.min(area.left, col),
            bottom: Math.max(area.bottom, row),
            right: Math.max(area.right, col)
          }
        : { top: row, left: col, bottom: row, right: col }
    }

    return area
  }

  merged(): string[] {
    const list = xml.child(this.root, X, 'mergeCells')
    return list ? xml.children(list, X, 'mergeCell').map((m) => xml.attr(m, 'ref')) : []
  }

  value(c: Element): CellValue {
    const t = c.getAttribute('t') || 'n'
    const v = xml.child(c, X, 'v')
    const raw = v ? xml.text(v) : null
    if (t === 's') {
      return raw === null || raw === '' ? null : this.host.strings.get(Number(raw))
    }

    if (t === 'inlineStr') {
      const is = xml.child(c, X, 'is')
      return is ? runText(is) : ''
    }

    if (raw === null || raw === '') {
      return null
    }

    if (t === 'b') {
      return raw === '1' || raw === 'true'
    }

    if (t === 'str' || t === 'e' || t === 'd') {
      return raw
    }

    const n = Number(raw)
    if (this.host.styles.isDate(Number(c.getAttribute('s') || 0))) {
      return serialToIso(n, this.host.date1904)
    }

    return n
  }

  formula(c: Element): string | null {
    const f = xml.child(c, X, 'f')
    if (!f) {
      return null
    }

    const text = xml.text(f)
    if (f.getAttribute('t') !== 'shared' || text) {
      return `=${text}`
    }

    const master = this.sharedMasters().get(xml.attr(f, 'si'))
    if (!master) {
      return null
    }

    const at = this.positionOf(c)

    return `=${formulas.shift(master.text, { rows: at.row - master.row, cols: at.col - master.col })}`
  }

  resetMasters(): void {
    this.masters = null
  }

  read(opts: ReadOptions = {}): SheetRead {
    const limit = opts.limit === undefined ? 10000 : opts.limit
    if (!(typeof limit === 'number' && Number.isInteger(limit) && limit > 0)) {
      throw new DocsError(
        'INVALID_RANGE',
        `limit is a positive number of cells, got ${String(opts.limit)}`
      )
    }

    const area = opts.range ? refs.parseRange(opts.range) : this.usedRange()
    const out: SheetRead = {
      name: this.name,
      range: null,
      rows: [],
      formulas: {},
      merged: this.merged()
    }
    if (!area) {
      return out
    }

    const width = area.right - area.left + 1
    let bottom = area.bottom
    if ((bottom - area.top + 1) * width > limit) {
      bottom = area.top + Math.max(1, Math.floor(limit / width)) - 1
      out.truncated = true
    }

    const window = { ...area, bottom }
    out.range = refs.range(window)
    out.rows = Array.from({ length: bottom - area.top + 1 }, (): CellValue[] =>
      new Array(width).fill(null)
    )
    this.fillRead(out, window)

    return out
  }

  get(ref: unknown): CellRead {
    const { row, col } = refs.parse(ref)
    const found = this.cells().find((c) => c.row === row && c.col === col)
    if (!found) {
      return { value: null, formula: null, format: 'General' }
    }

    return {
      value: this.value(found.el),
      formula: this.formula(found.el),
      format: this.host.styles.formatCode(Number(found.el.getAttribute('s') || 0))
    }
  }

  findRow(r: number): Element | null {
    return this.numberedRows().find((entry) => entry.r === r)?.el ?? null
  }

  ensureRow(r: number): Element {
    this.number()
    const rows = this.rows()
    const existing = rows.find((row) => rowNumber(row) - 1 === r)
    if (existing) {
      return existing
    }

    const el = xml.setAttributes(xml.named(this.doc, 'row', X), { r: r + 1 })
    const next = rows.find((row) => rowNumber(row) - 1 > r)
    this.sheetData().insertBefore(el, next ?? null)

    return el
  }

  findCell(pos: Position): Element | null {
    const row = this.findRow(pos.row)
    if (!row) {
      return null
    }

    return this.numberedCells(row, pos.row).find((entry) => entry.col === pos.col)?.el ?? null
  }

  ensureCell(pos: Position): Element {
    const row = this.ensureRow(pos.row)
    const cells = xml.children(row, X, 'c')
    const existing = cells.find((c) => refs.parse(c.getAttribute('r')).col === pos.col)
    if (existing) {
      return existing
    }

    const el = xml.setAttributes(xml.named(this.doc, 'c', X), { r: refs.cell(pos.row, pos.col) })
    const next = cells.find((c) => refs.parse(c.getAttribute('r')).col > pos.col)
    row.insertBefore(el, next ?? null)

    return el
  }

  write(ref: string, input: unknown): void {
    const pos = refs.parse(ref)
    this.guardWrite(pos, ref)
    const spec = normalize(input)
    const c = this.ensureCell(pos)
    if (spec.content) {
      this.writeContent(c, spec)
    }

    if (Object.keys(spec.style).length) {
      c.setAttribute(
        's',
        String(this.host.styles.derive(Number(c.getAttribute('s') || 0), spec.style))
      )
    }

    this.host.state.cellsChanged = true
    this.touch()
  }

  lastRow(): number {
    const used = this.usedRange()
    return used ? used.bottom : -1
  }

  shiftRows(from: number, delta: number): void {
    this.number()
    const rows = this.rows()
    const ordered = delta > 0 ? rows.slice().reverse() : rows
    for (const row of ordered) {
      const r = rowNumber(row) - 1
      if (r < from) {
        continue
      }

      row.setAttribute('r', String(r + delta + 1))
      for (const c of xml.children(row, X, 'c')) {
        c.setAttribute('r', refs.cell(r + delta, refs.parse(c.getAttribute('r')).col))
      }
    }

    this.host.state.cellsChanged = true
    this.touch()
  }

  removeRows(from: number, count: number): void {
    this.number()
    for (const row of this.rows()) {
      const r = rowNumber(row) - 1
      if (r >= from && r < from + count) {
        xml.remove(row)
      }
    }

    this.shiftRows(from + count, -count)
  }

  removeColumns(from: number, count: number): void {
    this.number()
    for (const row of this.rows()) {
      for (const c of xml.children(row, X, 'c')) {
        const pos = refs.parse(c.getAttribute('r'))
        if (pos.col >= from && pos.col < from + count) {
          xml.remove(c)
        } else if (pos.col >= from + count) {
          c.setAttribute('r', refs.cell(pos.row, pos.col - count))
        }
      }
    }

    const cols = xml.child(this.root, X, 'cols')
    if (cols) {
      this.shrinkColumns(cols, { at: from, count })
    }

    this.host.state.cellsChanged = true
    this.touch()
  }

  setColumnWidth(spec: ColumnWidth): void {
    const n = spec.col + 1
    const cols =
      xml.child(this.root, X, 'cols') ??
      xml.insertOrdered(this.root, xml.named(this.doc, 'cols', X), SHEET_ORDER)
    for (const el of xml.children(cols, X, 'col')) {
      columnSpan(el, n)
    }

    const el = xml.setAttributes(xml.named(this.doc, 'col', X), {
      min: n,
      max: n,
      width: spec.width,
      customWidth: 1
    })
    const next = xml.children(cols, X, 'col').find((c) => Number(c.getAttribute('min')) > n)
    cols.insertBefore(el, next ?? null)
    this.touch()
  }

  updateDimension(): void {
    const used = this.usedRange()
    const dim =
      xml.child(this.root, X, 'dimension') ??
      xml.insertOrdered(this.root, xml.named(this.doc, 'dimension', X), SHEET_ORDER)
    dim.setAttribute('ref', used ? refs.range(used) : 'A1')
  }

  private position(ref: string): Position {
    try {
      return refs.parse(ref)
    } catch {
      throw new DocsError(
        'BROKEN_XLSX',
        `sheet "${this.name}" has a cell with the invalid reference "${ref}"`
      )
    }
  }

  private sharedMasters(): Map<string, Master> {
    if (this.masters) return this.masters
    const out = new Map<string, Master>()
    for (const f of xml.descendants(this.doc, X, 'f')) {
      if (f.getAttribute('t') !== 'shared' || !f.hasAttribute('ref') || !xml.text(f)) {
        continue
      }

      const at = this.positionOf(parentCell(f))
      out.set(xml.attr(f, 'si'), { text: xml.text(f), row: at.row, col: at.col })
    }

    this.masters = out

    return out
  }

  private fillRead(out: SheetRead, area: Area): void {
    for (const { el, row, col } of this.cells()) {
      if (!within({ row, col }, area)) continue
      out.rows[row - area.top][col - area.left] = this.value(el)
      const formula = this.formula(el)
      if (formula) {
        out.formulas[refs.cell(row, col)] = formula
      }
    }
  }

  private appendValue(c: Element, value: string): void {
    xml
      .append(c, xml.named(this.doc, 'v', X))
      .appendChild(this.doc.createTextNode(xml.clean(value)))
  }

  private writeContent(c: Element, spec: CellSpec): void {
    for (const n of Array.from(c.childNodes)) {
      c.removeChild(n)
    }

    c.removeAttribute('t')
    if (spec.formula !== undefined) {
      const f = xml.append(c, xml.named(this.doc, 'f', X))
      f.appendChild(this.doc.createTextNode(xml.clean(spec.formula.replace(/^=/, ''))))
      this.host.state.formulas = true
      return
    }

    if (spec.date !== undefined) {
      const { serial, time } = isoToSerial(spec.date, this.host.date1904)
      this.appendValue(c, String(serial))
      if (spec.style.format === undefined) {
        spec.style.format = time ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd'
      }

      return
    }

    this.writeValue(c, spec.value ?? null)
  }

  private writeValue(c: Element, value: CellValue): void {
    if (typeof value === 'number') {
      this.appendValue(c, String(value))
    } else if (typeof value === 'boolean') {
      c.setAttribute('t', 'b')
      this.appendValue(c, value ? '1' : '0')
    } else if (typeof value === 'string') {
      c.setAttribute('t', 's')
      this.appendValue(c, String(this.host.strings.add(value)))
    }
  }

  private shrinkColumns(cols: Element, removed: Removal): void {
    const start = removed.at + 1
    const end = removed.at + removed.count
    for (const el of xml.children(cols, X, 'col')) {
      const min = Number(el.getAttribute('min'))
      const max = Number(el.getAttribute('max'))
      const newMin = min < start ? min : Math.max(start, min - removed.count)
      const newMax = shrunkEnd(max, { at: start, end, count: removed.count })
      if (newMax < newMin) {
        xml.remove(el)
      } else {
        xml.setAttributes(el, { min: newMin, max: newMax })
      }
    }

    if (!xml.children(cols, X, 'col').length) {
      xml.remove(cols)
    }
  }

  private guardMerged(pos: Position, ref: string): void {
    for (const range of this.merged()) {
      const m = refs.parseRange(range)
      if (within(pos, m) && (pos.row !== m.top || pos.col !== m.left)) {
        throw new DocsError(
          'CELL_MERGED',
          `${ref} is inside the merged range ${range}; write its top-left cell ${refs.cell(m.top, m.left)}`
        )
      }
    }
  }

  private guardWrite(pos: Position, ref: string): void {
    this.guardMerged(pos, ref)
    for (const f of xml.descendants(this.doc, X, 'f')) {
      const kind = f.getAttribute('t')
      const span = f.getAttribute('ref')
      if (!span || (kind !== 'shared' && kind !== 'array')) {
        continue
      }

      const area = refs.parseRange(span)
      if (area.top === area.bottom && area.left === area.right) {
        continue
      }

      const owner = this.positionOf(parentCell(f))
      const isOwner = owner.row === pos.row && owner.col === pos.col
      if ((kind === 'shared' && isOwner) || (kind === 'array' && within(pos, area))) {
        throw new DocsError(
          'SHARED_FORMULA',
          `${ref} is part of a ${kind} formula over ${span}; edit that range in Excel or rewrite every cell in it`
        )
      }
    }
  }
}
