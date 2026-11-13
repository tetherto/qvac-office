import * as xml from '../xml.ts'
import * as refs from './refs.ts'
import { parse, type Node } from './expression.ts'
import { dependencies, result } from './evaluate.ts'
import { Cells, Grid, type GridSheet, type NameEntry } from './grid.ts'
import { order } from './order.ts'
import { ErrorValue, Unsupported, errorNamed, isError, rounded } from './values.ts'
import { runText, type SharedStrings } from './strings.ts'
import type { Element } from '../xml.ts'
import type { EditState, Sheet } from './sheet.ts'
import type { Scalar } from './values.ts'

export type { NameEntry }

export interface Calculated {
  computed: number
  skipped: number
  cycles: string[]
  updated: number
}

export interface CalcHost {
  strings: SharedStrings
  worksheets(): Sheet[]
  names(): NameEntry[]
}

interface Formula {
  sheet: Sheet
  grid: GridSheet
  el: Element
  row: number
  col: number
  node: Node | null
  deps: number[]
}

interface Loaded {
  grid: Grid
  formulas: Formula[]
  index: Map<string, Cells<number>>
}

interface Encoded {
  type: string | null
  text: string
}

const X = xml.ns.x

export class CalcState implements EditState {
  touched: Set<Sheet>
  stale: boolean
  private changedCells: boolean
  private changedFormulas: boolean

  constructor() {
    this.touched = new Set()
    this.stale = false
    this.changedCells = false
    this.changedFormulas = false
  }

  get cellsChanged(): boolean {
    return this.changedCells
  }

  set cellsChanged(value: boolean) {
    this.changedCells = value
    this.stale ||= value
  }

  get formulas(): boolean {
    return this.changedFormulas
  }

  set formulas(value: boolean) {
    this.changedFormulas = value
    this.stale ||= value
  }
}

function errorValue(code: string): ErrorValue {
  try {
    return errorNamed(code)
  } catch {
    return new ErrorValue(code)
  }
}

function cached(el: Element, strings: SharedStrings): Scalar {
  const t = el.getAttribute('t') || 'n'
  const v = xml.child(el, X, 'v')
  const raw = v ? xml.text(v) : null
  if (t === 'inlineStr') {
    const is = xml.child(el, X, 'is')
    return is ? runText(is) : ''
  }

  if (raw === null || (raw === '' && t !== 'str')) {
    return null
  }

  if (t === 's') {
    return strings.get(Number(raw))
  }

  if (t === 'b') {
    return raw === '1' || raw === 'true'
  }

  if (t === 'e') {
    return errorValue(raw)
  }

  if (t === 'str' || t === 'd') {
    return raw
  }

  return Number(raw)
}

function formulaNode(sheet: Sheet, el: Element, f: Element): Node | null {
  const kind = f.getAttribute('t')
  if (
    kind === 'array' ||
    kind === 'dataTable' ||
    el.hasAttribute('cm') ||
    f.getAttribute('ca') === '1'
  ) {
    return null
  }

  const text = sheet.formula(el)
  if (!text) {
    return null
  }

  try {
    return parse(text)
  } catch (err) {
    if (err instanceof Unsupported) {
      return null
    }

    throw err
  }
}

function load(host: CalcHost): Loaded {
  const sheets: GridSheet[] = []
  const formulas: Formula[] = []
  const index = new Map<string, Cells<number>>()
  for (const sheet of host.worksheets()) {
    const grid: GridSheet = {
      key: sheet.name.toLowerCase(),
      index: sheet.index,
      cells: new Cells()
    }
    const own = new Cells<number>()
    sheets.push(grid)
    index.set(grid.key, own)
    for (const { el, row, col } of sheet.cells()) {
      grid.cells.set(row, col, cached(el, host.strings))
      const f = xml.child(el, X, 'f')
      if (!f) {
        continue
      }

      own.set(row, col, formulas.length)
      formulas.push({ sheet, grid, el, row, col, node: formulaNode(sheet, el, f), deps: [] })
    }
  }

  return { grid: new Grid(sheets, host.names()), formulas, index }
}

function link(loaded: Loaded): void {
  const { grid, formulas, index } = loaded
  for (const formula of formulas) {
    if (!formula.node) continue
    grid.at(formula.grid)
    try {
      for (const range of dependencies(formula.node, grid)) {
        index.get(range.sheet)?.each(range.area, (id) => formula.deps.push(id))
      }
    } catch (err) {
      if (!(err instanceof Unsupported)) {
        throw err
      }

      formula.node = null
    }
  }
}

function same(a: Scalar, b: Scalar): boolean {
  if (typeof a === 'number' && typeof b === 'number') {
    return rounded(a) === rounded(b)
  }

  if (isError(a) && isError(b)) {
    return a.code === b.code
  }

  return a === b
}

function encode(value: Scalar): Encoded {
  if (typeof value === 'string') {
    return { type: 'str', text: value }
  }

  if (typeof value === 'boolean') {
    return { type: 'b', text: value ? '1' : '0' }
  }

  if (isError(value)) {
    return { type: 'e', text: value.code }
  }

  return { type: null, text: String(value) }
}

function store(formula: Formula, value: Scalar, strings: SharedStrings): boolean {
  const { el, sheet } = formula
  if (same(cached(el, strings), value)) {
    return false
  }

  const { type, text } = encode(value)
  if (type) {
    el.setAttribute('t', type)
  } else {
    el.removeAttribute('t')
  }

  let v = xml.child(el, X, 'v')
  if (!v) {
    v = xml.named(sheet.doc, 'v', X)
    el.insertBefore(v, xml.child(el, X, 'f')?.nextSibling ?? null)
  }

  xml.setText(v, xml.clean(text))
  sheet.package.touch(sheet.part)

  return true
}

function compute(formula: Formula, grid: Grid): Scalar | undefined {
  if (!formula.node) return undefined
  grid.at(formula.grid)
  try {
    return result(formula.node, grid)
  } catch (err) {
    if (err instanceof Unsupported) {
      return undefined
    }

    throw err
  }
}

function label(formula: Formula): string {
  return `${refs.quoteSheet(formula.sheet.name)}!${refs.cell(formula.row, formula.col)}`
}

export function calculate(host: CalcHost): Calculated {
  const loaded = load(host)
  link(loaded)
  const { formulas, grid } = loaded
  const ordered = order(formulas.map((f) => f.deps))
  const known = new Uint8Array(formulas.length)
  const out: Calculated = { computed: 0, skipped: 0, cycles: [], updated: 0 }
  for (const id of ordered.order) {
    const formula = formulas[id]
    if (ordered.cyclic.has(id)) {
      out.cycles.push(label(formula))
      continue
    }

    const value = formula.deps.every((d) => known[d]) ? compute(formula, grid) : undefined
    if (value === undefined) {
      out.skipped++
      continue
    }

    known[id] = 1
    out.computed++
    formula.grid.cells.set(formula.row, formula.col, value)
    if (store(formula, value, host.strings)) {
      out.updated++
    }
  }

  return out
}
