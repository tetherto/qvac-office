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
  maxChars?: number
  detail?: 'runs'
}

export interface SlidePage {
  slides: SlideInfo[]
  next?: number
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
  if (isRecord(opts) && 'maxChars' in opts) {
    throw new DocsError(
      'INVALID_OPTIONS',
      'read returns every slide asked for; use readPage({ from, maxChars }) for a budget',
      {
        field: 'maxChars'
      }
    )
  }

  if (!Array.isArray(opts.slides)) {
    return opts
  }

  return { ...opts, slides: opts.slides.map((ref) => slides.slideIndex(deck, ref)) }
}

function checkBudget(maxChars: unknown): number {
  if (maxChars === undefined) {
    return Infinity
  }

  if (typeof maxChars !== 'number' || !Number.isInteger(maxChars) || maxChars < 1) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `maxChars must be a whole number above 0, got ${String(maxChars)}`,
      {
        field: 'maxChars',
        got: maxChars
      }
    )
  }

  return maxChars
}

export function readPage(deck: slides.Deck, opts: PageOptions, reader: SlideReader): SlidePage {
  checkDetail(opts.detail)
  const budget = checkBudget(opts.maxChars)
  const all = slides.entries(deck)
  const from = opts.from === undefined ? 0 : slides.slideIndex(deck, opts.from)
  const out: SlideInfo[] = []
  let used = 0
  for (let index = from; index < all.length; index++) {
    const info = reader(all[index], index)
    const size = JSON.stringify(info).length
    if (out.length > 0 && used + size > budget) {
      return { slides: out, next: index }
    }

    out.push(info)
    used += size
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
