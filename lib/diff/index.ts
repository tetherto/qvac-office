import { DocsError } from '../errors.ts'
import * as pptx from '../pptx/index.ts'
import * as docx from '../docx/index.ts'
import * as xlsx from '../xlsx/index.ts'
import * as refs from '../xlsx/refs.ts'
import { formatOf, open, type Doc, type Format } from '../tools/index.ts'
import { sequence, type Edit, type Entry } from './sequence.ts'

export type ChangeKind = 'added' | 'removed' | 'changed'

export interface Change {
  kind: ChangeKind
  where: string
  before?: string
  after?: string
}

export interface Counts {
  added: number
  removed: number
  changed: number
}

export interface Diff {
  format: Format
  changes: Change[]
  counts: Counts
  truncated: boolean
}

export interface DiffOptions {
  limit?: number
  excerpt?: number
}

type Input = Uint8Array | Doc

interface Compared {
  changes: Change[]
  truncated: boolean
}

interface SheetCells {
  values: Map<string, string>
  truncated: boolean
}

const LIMIT = 200
const EXCERPT = 120
const CELLS = 1000000

function kindOf<T>(edit: Edit<T>): ChangeKind {
  return edit.kind === 'same' ? 'changed' : edit.kind
}

function blockText(block: docx.Block): string {
  if (block.type === 'table') {
    return block.rows.map((row) => row.join(' | ')).join('\n')
  }

  return block.text
}

function partChanges(a: docx.ReadResult, b: docx.ReadResult): Change[] {
  return (['headers', 'footers'] as const)
    .filter((part) => a[part].join('\n') !== b[part].join('\n'))
    .map((part) => ({
      kind: 'changed',
      where: part,
      before: a[part].join('\n'),
      after: b[part].join('\n')
    }))
}

function docxChanges(before: docx.Document, after: docx.Document): Change[] {
  const a = before.read()
  const b = after.read()
  const blocks = sequence(a.blocks, b.blocks, (block) => `${block.type}:${blockText(block)}`)
    .filter((edit) => edit.kind !== 'same')
    .map((edit) => ({
      kind: kindOf(edit),
      where: `block ${edit.after?.index ?? edit.before?.index}`,
      before: edit.before && blockText(edit.before.item),
      after: edit.after && blockText(edit.after.item)
    }))

  return [...blocks, ...partChanges(a, b)]
}

function shapeText(shape: pptx.ShapeInfo): string {
  const text = (shape.paragraphs ?? []).map((p) => p.text).join('\n')
  const rows = (shape.rows ?? []).map((row) => row.join(' | ')).join('\n')
  const box = shape.box
    ? ` @${shape.box.x},${shape.box.y} ${shape.box.width}x${shape.box.height}`
    : ''

  return `${shape.kind} ${shape.name}${box}${text ? `: ${text}` : ''}${rows ? `: ${rows}` : ''}`
}

function slideKey(slide: pptx.SlideInfo): string {
  return 'id' in slide && typeof slide.id === 'number' ? String(slide.id) : `#${slide.index}`
}

function notesChange(index: number, a: pptx.SlideInfo, b: pptx.SlideInfo): Change[] {
  if ((a.notes ?? '') === (b.notes ?? '')) {
    return []
  }

  return [{ kind: 'changed', where: `slide ${index} notes`, before: a.notes, after: b.notes }]
}

function slideChanges(index: number, a: pptx.SlideInfo, b: pptx.SlideInfo): Change[] {
  const shapes = sequence(a.shapes, b.shapes, (shape) => String(shape.id ?? shape.name)).flatMap(
    (edit): Change[] => {
      const before = edit.before && shapeText(edit.before.item)
      const after = edit.after && shapeText(edit.after.item)
      const id = edit.after?.item.id ?? edit.before?.item.id
      if (edit.kind === 'same' && before === after) {
        return []
      }

      return [{ kind: kindOf(edit), where: `slide ${index} shape ${id}`, before, after }]
    }
  )

  return [...shapes, ...notesChange(index, a, b)]
}

function wholeSlide(edit: Edit<pptx.SlideInfo>): Change[] {
  const text = (entry?: Entry<pptx.SlideInfo>) => entry?.item.shapes.map(shapeText).join('\n')
  const index = edit.after?.index ?? edit.before?.index

  return [
    {
      kind: kindOf(edit),
      where: `slide ${index}`,
      before: text(edit.before),
      after: text(edit.after)
    }
  ]
}

function pptxChanges(before: pptx.Presentation, after: pptx.Presentation): Change[] {
  return sequence(before.read(), after.read(), slideKey).flatMap((edit) => {
    if (edit.before && edit.after) {
      return slideChanges(edit.after.index, edit.before.item, edit.after.item)
    }

    return wholeSlide(edit)
  })
}

function cells(book: xlsx.Workbook, sheet: number): SheetCells {
  const read = book.read(sheet, { limit: CELLS })
  const origin = refs.parse((read.range ?? 'A1').split(':')[0])
  const out = new Map<string, string>()
  read.rows.forEach((row, r) => {
    row.forEach((value, c) => {
      const ref = refs.cell(origin.row + r, origin.col + c)
      const formula = read.formulas[ref]
      if (value !== null || formula) {
        out.set(ref, formula ? `${formula} = ${String(value)}` : String(value))
      }
    })
  })

  return { values: out, truncated: read.truncated === true }
}

function cellKind(before: string | undefined, after: string | undefined): ChangeKind {
  if (before === undefined) {
    return 'added'
  }

  return after === undefined ? 'removed' : 'changed'
}

function sheetChanges(name: string, a: Map<string, string>, b: Map<string, string>): Change[] {
  return [...new Set([...a.keys(), ...b.keys()])].flatMap((ref): Change[] => {
    const before = a.get(ref)
    const after = b.get(ref)
    if (before === after) {
      return []
    }

    return [{ kind: cellKind(before, after), where: `${name}!${ref}`, before, after }]
  })
}

function worksheets(book: xlsx.Workbook): Map<string, number> {
  const sheets = book.sheets().filter((s) => s.kind === 'worksheet')

  return new Map(sheets.map((s) => [s.name, s.index]))
}

function xlsxChanges(before: xlsx.Workbook, after: xlsx.Workbook): Compared {
  const a = worksheets(before)
  const b = worksheets(after)
  let truncated = false
  const changes = [...new Set([...a.keys(), ...b.keys()])].flatMap((name): Change[] => {
    const i = a.get(name)
    const j = b.get(name)
    if (i === undefined || j === undefined) {
      return [{ kind: i === undefined ? 'added' : 'removed', where: `sheet ${name}` }]
    }

    const x = cells(before, i)
    const y = cells(after, j)
    truncated ||= x.truncated || y.truncated

    return sheetChanges(name, x.values, y.values)
  })

  return { changes, truncated }
}

function changesOf(a: Doc, b: Doc): Compared {
  if (a instanceof docx.Document && b instanceof docx.Document) {
    return { changes: docxChanges(a, b), truncated: false }
  }

  if (a instanceof pptx.Presentation && b instanceof pptx.Presentation) {
    return { changes: pptxChanges(a, b), truncated: false }
  }

  if (a instanceof xlsx.Workbook && b instanceof xlsx.Workbook) {
    return xlsxChanges(a, b)
  }

  throw new DocsError(
    'INVALID_INPUT',
    `diff compares two documents of one format, got ${formatOf(a)} and ${formatOf(b)}`
  )
}

function clip(text: string | undefined, length: number): string | undefined {
  return text === undefined || text.length <= length ? text : `${text.slice(0, length - 1)}…`
}

function load(input: Input): Doc {
  return input instanceof Uint8Array ? open(input) : input
}

export function diff(before: Input, after: Input, opts: DiffOptions = {}): Diff {
  const a = load(before)
  const { changes: all, truncated } = changesOf(a, load(after))
  const limit = opts.limit ?? LIMIT
  const excerpt = opts.excerpt ?? EXCERPT
  const changes = all.slice(0, limit).map((change) => ({
    ...change,
    before: clip(change.before, excerpt),
    after: clip(change.after, excerpt)
  }))
  const counts = { added: 0, removed: 0, changed: 0 }
  for (const change of all) {
    counts[change.kind]++
  }

  return { format: formatOf(a), changes, counts, truncated: truncated || all.length > limit }
}
