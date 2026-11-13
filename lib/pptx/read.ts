import * as xml from '../xml.ts'
import * as text from '../drawingml/text.ts'
import { DocsError } from '../errors.ts'
import type { Document } from '../xml.ts'
import type { ShapeEntry } from './shapes.ts'
import { shapesOf } from './shapes.ts'
import type { ReadOptions, ShapeInfo } from './types.ts'

const { a: A, p: P } = xml.ns
export const BODY_TYPES = new Set(['body', 'obj', 'subTitle'])

interface Window {
  from: number
  to: number
}

function window(total: number, opts: ReadOptions): Window {
  const from = opts.from ?? 0
  const count = opts.count ?? Math.max(total, 1)
  if (!Number.isInteger(from) || from < 0 || (from >= total && total > 0)) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `from must be a slide index from 0 to ${Math.max(0, total - 1)}, got ${String(opts.from)}`
    )
  }

  if (!Number.isInteger(count) || count < 1) {
    throw new DocsError('INVALID_OPTIONS', `count must be at least 1, got ${String(opts.count)}`)
  }

  return { from, to: Math.min(total, from + count) }
}

export function pick(total: number, opts: ReadOptions): (i: number) => boolean {
  if (opts.slides === undefined) {
    const { from, to } = window(total, opts)
    return (i) => i >= from && i < to
  }

  if (opts.from !== undefined || opts.count !== undefined) {
    throw new DocsError('INVALID_OPTIONS', 'pass either slides or from/count, not both')
  }

  const slides: unknown = opts.slides
  const valid =
    Array.isArray(slides) && slides.every((i) => Number.isInteger(i) && i >= 0 && i < total)
  if (!valid) {
    throw new DocsError('SLIDE_NOT_FOUND', `slides must list slide indexes from 0 to ${total - 1}`)
  }

  const set = new Set<unknown>(slides)

  return (i) => set.has(i)
}

export function layoutName(doc: Document): string {
  const cSld = xml.child(xml.root(doc), P, 'cSld')
  return cSld?.getAttribute('name') || ''
}

export function layoutType(doc: Document): string {
  return xml.root(doc).getAttribute('type') || 'cust'
}

export function hasBody(doc: Document): boolean {
  return shapesOf(doc).some(
    (s) => s.placeholder !== null && BODY_TYPES.has(s.placeholder) && s.placeholder !== 'subTitle'
  )
}

function tableRows(shape: ShapeEntry): string[][] {
  return xml.descendants(shape.el, A, 'tr').map((tr) =>
    xml.children(tr, A, 'tc').map((tc) => {
      const txBody = xml.child(tc, A, 'txBody')
      return txBody ? text.readText(txBody) : ''
    })
  )
}

export function describe(shape: ShapeEntry): ShapeInfo {
  const out: ShapeInfo = { id: shape.id, name: shape.name, kind: shape.kind }
  if (shape.placeholder) {
    out.placeholder = shape.placeholder
  }

  if (shape.group !== null) {
    out.group = shape.group
  }

  if (shape.kind === 'text') {
    const txBody = xml.child(shape.el, P, 'txBody')
    out.paragraphs = txBody ? text.readParagraphs(txBody) : []
  } else if (shape.kind === 'table') {
    out.rows = tableRows(shape)
  }

  return out
}
