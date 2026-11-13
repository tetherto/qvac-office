import * as refs from './refs.ts'

export type RefKind = 'cell' | 'range' | 'columns' | 'rows'

export interface Point {
  col?: number
  row?: number
  absCol?: boolean
  absRow?: boolean
}

export interface Ref {
  kind: RefKind
  start: Point
  end: Point | null
  length: number
}

export interface Change {
  sheet: string
  axis: 'row' | 'col'
  at: number
  count: number
  kind: 'insert' | 'delete'
}

export interface RefContext {
  sheet: string | null
  original: string
}

export interface Handlers {
  onRef?: (ref: Ref, context: RefContext) => string | null
  onSheet?: (names: string[], text: string) => string
  onTable?: (name: string) => void
}

export interface Rename {
  from: string
  to: string
}

export interface Offset {
  rows: number
  cols: number
}

interface ScanState {
  formula: string
  at: number
  sheet: string | null
  external: boolean
  out: string
}

const NAME_START = /[A-Za-z_À-￿]/
const WORD = /[A-Za-z0-9_.À-￿]/
const BARE_PREFIX = /[A-Za-z_À-￿][\w.À-￿]*(?::[A-Za-z_À-￿][\w.À-￿]*)?!/y
const CELL_RANGE =
  /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?(?![A-Za-z0-9_(![])/y
const COLUMN_RANGE = /(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(![])/y
const ROW_RANGE = /(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![A-Za-z0-9_(.![])/y
const WORKBOOK = /\[[^\]]*\]/

function quoted(text: string, at: number, quote: string): number {
  let i = at + 1
  while (i < text.length) {
    if (text[i] === quote && text[i + 1] === quote) {
      i += 2
    } else if (text[i] === quote) {
      return i + 1
    } else {
      i++
    }
  }

  return text.length
}

function matchAt(re: RegExp, text: string, at: number): RegExpExecArray | null {
  re.lastIndex = at
  return re.exec(text)
}

function inColumns(letters: string | undefined): boolean {
  return letters === undefined || refs.columnIndex(letters) < refs.MAX_COLUMNS
}

function cellRef(m: RegExpExecArray): Ref {
  const start = {
    col: refs.columnIndex(m[2]),
    row: Number(m[4]) - 1,
    absCol: m[1] === '$',
    absRow: m[3] === '$'
  }
  const end = m[6]
    ? {
        col: refs.columnIndex(m[6]),
        row: Number(m[8]) - 1,
        absCol: m[5] === '$',
        absRow: m[7] === '$'
      }
    : null

  return { kind: end ? 'range' : 'cell', start, end, length: m[0].length }
}

function readRef(text: string, at: number): Ref | null {
  const cells = matchAt(CELL_RANGE, text, at)
  if (cells && inColumns(cells[2]) && inColumns(cells[6])) {
    return cellRef(cells)
  }

  const columns = matchAt(COLUMN_RANGE, text, at)
  if (columns && inColumns(columns[2]) && inColumns(columns[4])) {
    return {
      kind: 'columns',
      start: { col: refs.columnIndex(columns[2]), absCol: columns[1] === '$' },
      end: { col: refs.columnIndex(columns[4]), absCol: columns[3] === '$' },
      length: columns[0].length
    }
  }

  const rows = matchAt(ROW_RANGE, text, at)
  if (!rows) {
    return null
  }

  return {
    kind: 'rows',
    start: { row: Number(rows[2]) - 1, absRow: rows[1] === '$' },
    end: { row: Number(rows[4]) - 1, absRow: rows[3] === '$' },
    length: rows[0].length
  }
}

function bracketEnd(text: string, at: number): number {
  let depth = 0
  for (let i = at; i < text.length; i++) {
    const ch = text[i]
    if (ch === "'") {
      i++
    } else if (ch === '[') {
      depth++
    } else if (ch === ']' && --depth === 0) {
      return i + 1
    }
  }

  return text.length
}

function emitExternal(state: ScanState, text: string): void {
  state.out += text
  state.at += text.length
  state.sheet = text
  state.external = true
}

function emitSheet(state: ScanState, handlers: Handlers, prefix: string): void {
  const name = prefix.startsWith("'")
    ? prefix.slice(1, -2).replace(/''/g, "'")
    : prefix.slice(0, -1)
  if (WORKBOOK.test(name)) {
    emitExternal(state, prefix)
    return
  }

  const names = name.split(':')
  state.sheet = name
  state.external = false
  state.out += handlers.onSheet ? handlers.onSheet(names, prefix) : prefix
  state.at += prefix.length
}

function bracketStep(state: ScanState, handlers: Handlers): boolean {
  const { formula, at } = state
  if (formula[at] !== '[') {
    return false
  }

  const end = bracketEnd(formula, at)
  let start = at
  while (start > 0 && WORD.test(formula[start - 1])) {
    start--
  }

  const prefix = start === at ? matchAt(BARE_PREFIX, formula, end) : null
  if (prefix || (start === at && formula[end] === '!')) {
    emitExternal(state, formula.slice(at, end) + (prefix ? prefix[0] : '!'))
    return true
  }

  if (start < at && !state.external) {
    handlers.onTable?.(formula.slice(start, at))
  }

  state.out += formula.slice(at, end)
  state.at = end
  state.sheet = null
  state.external = false

  return true
}

function quotedStep(state: ScanState, handlers: Handlers): boolean {
  const { formula, at } = state
  const ch = formula[at]
  if (ch !== '"' && ch !== "'") {
    return false
  }

  const end = quoted(formula, at, ch)
  if (ch === "'" && formula[end] === '!') {
    emitSheet(state, handlers, formula.slice(at, end + 1))
    return true
  }

  state.out += formula.slice(at, end)
  state.at = end
  if (ch === '"') {
    state.sheet = null
  }

  return true
}

function referenceStep(state: ScanState, handlers: Handlers): boolean {
  const { formula, at } = state
  const prev = at > 0 ? formula[at - 1] : ''
  if (WORD.test(prev) || prev === '$') {
    return false
  }

  const bare =
    state.sheet === null && NAME_START.test(formula[at]) ? matchAt(BARE_PREFIX, formula, at) : null
  if (bare) {
    emitSheet(state, handlers, bare[0])
    return true
  }

  const ref = readRef(formula, at)
  if (!ref) {
    return false
  }

  const original = formula.slice(at, at + ref.length)
  const next =
    handlers.onRef && !state.external ? handlers.onRef(ref, { sheet: state.sheet, original }) : null
  state.out += next ?? original
  state.at += ref.length
  state.sheet = null
  state.external = false

  return true
}

function plainStep(state: ScanState): void {
  const ch = state.formula[state.at]
  if (state.sheet !== null && !WORD.test(ch)) {
    state.sheet = null
    state.external = false
  }

  state.out += ch
  state.at++
}

export function scan(formula: string, handlers: Handlers = {}): string {
  const state: ScanState = { formula, at: 0, sheet: null, external: false, out: '' }
  while (state.at < formula.length) {
    if (
      !bracketStep(state, handlers) &&
      !quotedStep(state, handlers) &&
      !referenceStep(state, handlers)
    ) {
      plainStep(state)
    }
  }

  return state.out
}

export function sheetsIn(formula: string): string[] {
  const found: string[] = []
  scan(formula, {
    onSheet: (names, text) => {
      found.push(...names)
      return text
    }
  })

  return found
}

export function tablesIn(formula: string): string[] {
  const found: string[] = []
  scan(formula, { onTable: (name) => found.push(name) })

  return found
}

export function usesTable(formula: string, table: string): boolean {
  const wanted = table.toLowerCase()
  return tablesIn(formula).some((name) => name.toLowerCase() === wanted)
}

export function mentions(formula: string, sheet: string): boolean {
  const wanted = sheet.toLowerCase()
  return sheetsIn(formula).some((name) => name.toLowerCase() === wanted)
}

export function renameSheet(formula: string, rename: Rename): string {
  const wanted = rename.from.toLowerCase()
  return scan(formula, {
    onSheet: (names, text) => {
      if (!names.some((name) => name.toLowerCase() === wanted)) {
        return text
      }

      const next = names.map((name) => (name.toLowerCase() === wanted ? rename.to : name))

      return `${quoteSheets(next)}!`
    }
  })
}

function quoteSheets(names: string[]): string {
  const quoted = names.map((name) => refs.quoteSheet(name))
  if (quoted.every((name, i) => name === names[i])) {
    return names.join(':')
  }

  return `'${names.map((name) => name.replace(/'/g, "''")).join(':')}'`
}

function formatCell(p: Point): string {
  return `${p.absCol ? '$' : ''}${refs.column(p.col ?? 0)}${p.absRow ? '$' : ''}${(p.row ?? 0) + 1}`
}

function formatRef(ref: Ref): string {
  const { start } = ref
  const end = ref.end ?? ref.start
  if (ref.kind === 'cell') {
    return formatCell(start)
  }

  if (ref.kind === 'range') {
    return `${formatCell(start)}:${formatCell(end)}`
  }

  if (ref.kind === 'columns') {
    return `${start.absCol ? '$' : ''}${refs.column(start.col ?? 0)}:${end.absCol ? '$' : ''}${refs.column(end.col ?? 0)}`
  }

  return `${start.absRow ? '$' : ''}${(start.row ?? 0) + 1}:${end.absRow ? '$' : ''}${(end.row ?? 0) + 1}`
}

function outside(p: Point): boolean {
  const badCol = p.col !== undefined && (p.col < 0 || p.col >= refs.MAX_COLUMNS)
  const badRow = p.row !== undefined && (p.row < 0 || p.row >= refs.MAX_ROWS)
  return badCol || badRow
}

const BROKEN = '\u0000'

function moved(p: Point, offset: Offset): Point {
  const col = p.col !== undefined && !p.absCol ? p.col + offset.cols : p.col
  const row = p.row !== undefined && !p.absRow ? p.row + offset.rows : p.row
  return { ...p, col, row }
}

export function shift(formula: string, offset: Offset): string {
  const out = scan(formula, {
    onRef: (ref) => {
      const start = moved(ref.start, offset)
      const end = ref.end ? moved(ref.end, offset) : null
      const broken = outside(start) || (end !== null && outside(end))
      return broken ? BROKEN : formatRef({ ...ref, start, end })
    }
  })
  return out.includes(BROKEN) ? '#REF!' : out
}

function limit(change: Change): number {
  return change.axis === 'row' ? refs.MAX_ROWS : refs.MAX_COLUMNS
}

export function moveIndex(i: number, change: Change): number | null {
  const { at, count, kind } = change
  if (kind === 'insert') {
    if (i < at) {
      return i
    }

    return i + count < limit(change) ? i + count : null
  }

  if (i < at) {
    return i
  }

  if (i >= at + count) {
    return i - count
  }

  return null
}

function spanStart(a: number, change: Change): number {
  if (a < change.at) return a
  return a >= change.at + change.count ? a - change.count : change.at
}

function spanEnd(b: number, change: Change): number {
  if (b < change.at) return b
  return b >= change.at + change.count ? b - change.count : change.at - 1
}

export function moveSpan(span: [number, number], change: Change): [number, number] | null {
  const [a, b] = span
  const { at, count } = change
  if (change.kind === 'insert') {
    const last = limit(change) - 1
    const first = a >= at ? a + count : a

    return first > last ? null : [first, b >= at ? Math.min(b + count, last) : b]
  }

  const first = spanStart(a, change)
  const last = spanEnd(b, change)

  return last < first ? null : [first, last]
}

function adjustRef(ref: Ref, change: Change): Ref | null {
  const key = change.axis === 'row' ? 'row' : 'col'
  const from = ref.start[key]
  if (from === undefined) {
    return ref
  }

  if (ref.kind === 'cell') {
    const moved = moveIndex(from, change)
    return moved === null ? null : { ...ref, start: { ...ref.start, [key]: moved } }
  }

  const to = ref.end?.[key]
  if (!ref.end || to === undefined) {
    return ref
  }

  const span = moveSpan([from, to], change)
  if (span === null) {
    return null
  }

  return { ...ref, start: { ...ref.start, [key]: span[0] }, end: { ...ref.end, [key]: span[1] } }
}

export function adjust(formula: string, ownSheet: string | null, change: Change): string {
  const target = change.sheet.toLowerCase()
  const own = ownSheet ? ownSheet.toLowerCase() : null
  return scan(formula, {
    onRef: (ref, { sheet, original }) => {
      const pointsAt = sheet !== null ? sheet.toLowerCase() : own
      if (pointsAt !== target) {
        return original
      }

      const moved = adjustRef(ref, change)

      return moved === null ? '#REF!' : formatRef(moved)
    }
  })
}

export function references(formula: string, ownSheet: string | null, sheet: string): boolean {
  const target = sheet.toLowerCase()
  const own = ownSheet ? ownSheet.toLowerCase() : null
  let hit = false
  scan(formula, {
    onRef: (ref, { sheet: prefix, original }) => {
      if ((prefix !== null ? prefix.toLowerCase() : own) === target) {
        hit = true
      }

      return original
    }
  })

  return hit
}

function plain(p: Point): Point {
  return { ...p, absCol: false, absRow: false }
}

export function adjustRange(range: string, change: Change): string | null {
  const ref = readRef(range, 0)
  if (!ref || ref.length !== range.length) {
    return range
  }

  const moved = adjustRef(ref, change)
  if (moved === null) {
    return null
  }

  return formatRef({
    ...moved,
    start: plain(moved.start),
    end: moved.end ? plain(moved.end) : null
  })
}
