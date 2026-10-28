import * as xml from '../xml.ts'
import * as text from '../drawingml/text.ts'
import * as slides from './slides.ts'
import * as notes from './notes.ts'
import { DocsError } from '../errors.ts'
import { isRecord, errorMessage } from '../guards.ts'
import { escapeRegex } from '../drawingml/ranges.ts'
import { layoutName, BODY_TYPES } from './read.ts'
import { shapesOf, type ShapeEntry, type ShapeKind } from './shapes.ts'
import type { Deck, SlideEntry } from './slides.ts'

export interface FindOptions {
  text?: string
  placeholder?: string
  kind?: ShapeKind
  name?: string
  regex?: boolean
  matchCase?: boolean
}

export interface FoundShape {
  slide: number
  slideId: number
  shape: number
  name: string
  kind: ShapeKind
  text: string
}

export interface SlideOutline {
  index: number
  id: number
  layout: string | null
  title: string | null
  shapes: number
  notes: boolean
  charts: number
  tables: number
  pictures: number
}

interface Slide {
  entry: SlideEntry
  index: number
}

type Test = (shape: ShapeEntry, value: string) => boolean

const { a: A, p: P } = xml.ns
const FIND_KEYS = ['text', 'placeholder', 'kind', 'name', 'regex', 'matchCase']
const KINDS: readonly ShapeKind[] = [
  'text',
  'shape',
  'picture',
  'group',
  'connector',
  'table',
  'chart',
  'diagram',
  'frame'
]
const ROLES: Record<string, (type: string) => boolean> = {
  title: (type) => type === 'title' || type === 'ctrTitle',
  subtitle: (type) => type === 'subTitle',
  body: (type) => BODY_TYPES.has(type) && type !== 'subTitle'
}

export function shapeText(shape: ShapeEntry): string {
  if (shape.kind === 'table') {
    return xml
      .descendants(shape.el, A, 'tc')
      .map((tc) => {
        const txBody = xml.child(tc, A, 'txBody')
        return txBody ? text.readText(txBody) : ''
      })
      .join('\n')
  }

  const txBody = shape.el.localName === 'sp' ? xml.child(shape.el, P, 'txBody') : null

  return txBody ? text.readText(txBody) : ''
}

function compile(source: string, opts: FindOptions, field: string): RegExp {
  try {
    return new RegExp(opts.regex ? source : escapeRegex(source), opts.matchCase ? 'u' : 'iu')
  } catch (err) {
    throw new DocsError('INVALID_OPTIONS', `${field} is not a valid regular expression`, {
      field,
      got: source,
      cause: errorMessage(err)
    })
  }
}

function checkKeys(opts: Record<string, unknown>): void {
  const unknownKey = Object.keys(opts).find((key) => !FIND_KEYS.includes(key))
  if (unknownKey) {
    throw new DocsError('INVALID_OPTIONS', `unknown find option "${unknownKey}"`, {
      field: unknownKey,
      candidates: FIND_KEYS
    })
  }

  if (opts.kind !== undefined && !KINDS.some((kind) => kind === opts.kind)) {
    throw new DocsError('INVALID_OPTIONS', `kind must be one of ${KINDS.join(', ')}`, {
      field: 'kind',
      got: opts.kind,
      candidates: KINDS
    })
  }

  for (const key of ['text', 'placeholder', 'name']) {
    if (opts[key] !== undefined && typeof opts[key] !== 'string') {
      throw new DocsError('INVALID_OPTIONS', `${key} must be a string`, { field: key })
    }
  }

  for (const key of ['regex', 'matchCase']) {
    if (opts[key] !== undefined && typeof opts[key] !== 'boolean') {
      throw new DocsError('INVALID_OPTIONS', `${key} must be true or false`, { field: key })
    }
  }
}

function placeholderTest(wanted: string): Test {
  const role = ROLES[wanted.toLowerCase()]
  return (shape) =>
    shape.placeholder !== null &&
    (role ? role(shape.placeholder) : shape.placeholder.toLowerCase() === wanted.toLowerCase())
}

function tests(opts: FindOptions): Test[] {
  const out: Test[] = []
  if (opts.kind !== undefined) {
    out.push((shape) => shape.kind === opts.kind)
  }

  if (opts.placeholder !== undefined) {
    out.push(placeholderTest(opts.placeholder))
  }

  if (opts.name !== undefined) {
    const name = compile(opts.name, opts, 'name')
    out.push((shape) => name.test(shape.name))
  }

  if (opts.text !== undefined) {
    const wanted = compile(opts.text, opts, 'text')
    out.push((_, value) => wanted.test(value))
  }

  return out
}

function slideMatches(deck: Deck, slide: Slide, all: Test[]): FoundShape[] {
  const { entry, index } = slide
  return shapesOf(deck.pkg.xml(entry.part)).flatMap((shape) => {
    const value = shapeText(shape)
    if (shape.id === null || !all.every((test) => test(shape, value))) {
      return []
    }

    return [
      {
        slide: index,
        slideId: Number(entry.id),
        shape: shape.id,
        name: shape.name,
        kind: shape.kind,
        text: value
      }
    ]
  })
}

export function find(deck: Deck, input: unknown = {}): FoundShape[] {
  if (!isRecord(input)) {
    throw new DocsError('INVALID_OPTIONS', 'find takes { text, placeholder, kind, name }')
  }

  checkKeys(input)
  const all = tests(input)

  return slides.entries(deck).flatMap((entry, index) => slideMatches(deck, { entry, index }, all))
}

function titleOf(shapes: ShapeEntry[]): string | null {
  const title = shapes.find((s) => s.placeholder !== null && ROLES.title(s.placeholder))
  const value = title ? shapeText(title).trim() : ''
  return value || null
}

function slideOutline(deck: Deck, entry: SlideEntry, index: number): SlideOutline {
  const shapes = shapesOf(deck.pkg.xml(entry.part))
  const layout = slides.layoutOf(deck.pkg, entry.part)
  const count = (kind: ShapeKind): number => shapes.filter((s) => s.kind === kind).length

  return {
    index,
    id: Number(entry.id),
    layout: layout ? layoutName(deck.pkg.xml(layout)) : null,
    title: titleOf(shapes),
    shapes: shapes.length,
    notes: notes.notesText(deck, entry.part) !== null,
    charts: count('chart'),
    tables: count('table'),
    pictures: count('picture')
  }
}

export function outline(deck: Deck): SlideOutline[] {
  return slides.entries(deck).map((entry, index) => slideOutline(deck, entry, index))
}
