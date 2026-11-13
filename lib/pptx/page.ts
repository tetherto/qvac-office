import * as xml from '../xml.ts'
import * as slides from './slides.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import { readRuns } from '../drawingml/runs.ts'
import type { Package } from '../opc/package.ts'
import type { ShapeEntry } from './shapes.ts'
import type { ReadOptions, ShapeInfo, SlideInfo, SlideRef } from './types.ts'

export interface PageOptions {
  from?: SlideRef
  limit?: number
  detail?: 'runs'
}

export interface SlidePage {
  slides: SlideInfo[]
  next?: { id: number }
  truncated?: true
}

export interface RunSource {
  pkg: Package
  part: string
}

export type SlideReader = (entry: slides.SlideEntry, index: number) => SlideInfo

const { a: A, p: P } = xml.ns

function checkDetail(detail: unknown): void {
  if (detail !== undefined && detail !== 'runs') {
    throw new DocsError(
      'INVALID_OPTIONS',
      `detail must be "runs" or left out, got ${String(detail)}`,
      {
        field: 'detail',
        candidates: ['runs']
      }
    )
  }
}

export function readOptions(deck: slides.Deck, opts: ReadOptions): ReadOptions {
  checkDetail(opts.detail)
  const budgeted = isRecord(opts) && ['limit', 'maxChars'].find((field) => field in opts)
  if (budgeted) {
    throw new DocsError(
      'INVALID_OPTIONS',
      'read returns every slide asked for; use readPage({ from, limit }) for a budget',
      {
        field: budgeted
      }
    )
  }

  if (!Array.isArray(opts.slides)) {
    return opts
  }

  return { ...opts, slides: opts.slides.map((ref) => slides.slideIndex(deck, ref)) }
}

function checkLimit(limit: unknown): number {
  if (limit === undefined) {
    return Infinity
  }

  if (typeof limit !== 'number' || !Number.isInteger(limit) || limit < 1) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `limit is a positive number of characters, got ${String(limit)}`,
      {
        field: 'limit',
        expected: 'positive integer',
        got: limit
      }
    )
  }

  return limit
}

function slideChars(slide: SlideInfo): number {
  const shapes = slide.shapes.reduce(
    (sum, shape) =>
      sum +
      (shape.paragraphs ?? []).reduce((chars, p) => chars + p.text.length, 0) +
      (shape.rows ?? []).reduce((chars, row) => chars + row.join('').length, 0),
    0
  )
  return shapes + (slide.notes ?? '').length
}

export function readPage(deck: slides.Deck, opts: PageOptions, reader: SlideReader): SlidePage {
  checkDetail(opts.detail)
  const limit = checkLimit(opts.limit)
  const all = slides.entries(deck)
  const from = opts.from === undefined ? 0 : slides.slideIndex(deck, opts.from)
  const out: SlideInfo[] = []
  let used = 0
  for (let index = from; index < all.length; index++) {
    const info = reader(all[index], index)
    used += slideChars(info)
    if (out.length > 0 && used > limit) {
      return { slides: out, next: { id: Number(all[index].id) }, truncated: true }
    }

    out.push(info)
  }

  return { slides: out }
}

export function addRuns(source: RunSource, shape: ShapeEntry, info: ShapeInfo): void {
  const { pkg, part } = source
  const txBody = shape.kind === 'text' ? xml.child(shape.el, P, 'txBody') : null
  if (!txBody || !info.paragraphs) {
    return
  }

  const rels = pkg.rels(part)
  const target = (rId: string): string | null => rels.get(rId)?.target ?? null
  const paragraphs = info.paragraphs
  xml.children(txBody, A, 'p').forEach((p, i) => {
    paragraphs[i].runs = readRuns(p, target)
  })
}
