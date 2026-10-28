import * as xml from '../xml.ts'
import * as ranges from '../drawingml/ranges.ts'
import * as slides from './slides.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import { shapesOf, type ShapeEntry } from './shapes.ts'
import type { Element } from '../xml.ts'
import type { Search, Hit } from '../drawingml/ranges.ts'
import type { Style } from '../drawingml/text.ts'
import type { ShapeId, ShapeRef, SlideRef } from './types.ts'

export interface ReplaceTextOptions extends Search {
  replace: string
  slide?: SlideRef
  shape?: ShapeId
  notes?: boolean
}

export interface TextShape {
  slide: number
  shape: number | 'notes'
}

export interface ReplacedText {
  replaced: number
  shapes: TextShape[]
  refused?: number
}

export interface FormattedText {
  formatted: number
  runs: number
  refused?: number
}

export interface LinkedText {
  linked: number
  runs: number
  refused?: number
}

export interface Located {
  entry: slides.SlideEntry
  found: ShapeEntry
}

export interface DeckText {
  deck: slides.Deck
  locate: (target: unknown) => Located
  touch: (part: string) => void
}

export interface FormatRequest {
  target: ShapeRef
  search: Search
  style: Style
}

export interface LinkRequest {
  target: ShapeRef
  search: Search
  url: string
}

interface Refused {
  refused?: number
}

interface Body {
  slide: number
  shape: number | 'notes'
  part: string
  txBody: Element
}

interface Matched {
  hits: Map<Body, Hit[]>
  count: number
  refused: number
}

const { a: A, p: P } = xml.ns
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
const HYPERLINK = REL_BASE + 'hyperlink'
const NOTES = REL_BASE + 'notesSlide'
const URL_PATTERN = /^(https?:\/\/[^\s/?#]+[^\s]*|mailto:[^\s@]+@[^\s@]+)$/i

function shapeBodies(found: ShapeEntry): Element[] {
  if (found.kind === 'table') {
    return xml.descendants(found.el, A, 'txBody')
  }

  const own = found.el.localName === 'sp' ? xml.child(found.el, P, 'txBody') : null

  return own ? [own] : []
}

function notesBody(ctx: DeckText, slidePart: string): Body[] {
  const part = ctx.deck.pkg.related(slidePart, NOTES)[0]
  if (!part) {
    return []
  }

  const body = shapesOf(ctx.deck.pkg.xml(part)).find((s) => s.placeholder === 'body')
  const txBody = body ? xml.child(body.el, P, 'txBody') : null

  return txBody ? [{ slide: -1, shape: 'notes', part, txBody }] : []
}

function slideBodies(ctx: DeckText, index: number, notes: boolean): Body[] {
  const entry = slides.entries(ctx.deck)[index]
  const shapes = shapesOf(ctx.deck.pkg.xml(entry.part)).flatMap((found) =>
    shapeBodies(found).map((txBody) => ({
      slide: index,
      shape: found.id ?? 0,
      part: entry.part,
      txBody
    }))
  )
  const extra = notes ? notesBody(ctx, entry.part).map((b) => ({ ...b, slide: index })) : []

  return [...shapes, ...extra]
}

function scopeFlag(value: unknown): boolean {
  if (value !== undefined && typeof value !== 'boolean') {
    throw new DocsError('INVALID_OPTIONS', 'notes must be true or false', { field: 'notes' })
  }

  return value === true
}

function scopedBodies(ctx: DeckText, opts: Record<string, unknown>): Body[] {
  const notes = scopeFlag(opts.notes)
  if (opts.shape !== undefined) {
    if (opts.slide === undefined || notes) {
      throw new DocsError('INVALID_OPTIONS', 'shape needs slide, and cannot be used with notes', {
        field: 'shape'
      })
    }

    return targetBodies(ctx, { slide: opts.slide, shape: opts.shape })
  }

  if (opts.slide !== undefined) {
    return slideBodies(ctx, slides.slideIndex(ctx.deck, opts.slide), notes)
  }

  return slides.entries(ctx.deck).flatMap((_, i) => slideBodies(ctx, i, notes))
}

function targetBodies(ctx: DeckText, target: unknown): Body[] {
  const { entry, found } = ctx.locate(target)
  const bodies = shapeBodies(found)
  if (bodies.length === 0) {
    throw new DocsError('NOT_TEXT', `shape ${found.id} is a ${found.kind} and holds no text`)
  }

  const slide = slides.entries(ctx.deck).findIndex((e) => e.part === entry.part)

  return bodies.map((txBody) => ({ slide, shape: found.id ?? 0, part: entry.part, txBody }))
}

function match(bodies: Body[], wanted: ranges.Pattern): Matched {
  const out: Matched = { hits: new Map(), count: 0, refused: 0 }
  const limit = wanted.all ? Infinity : 1
  for (const body of bodies) {
    const found = ranges.search(body.txBody, wanted, limit - out.count)
    out.refused += found.refused
    if (found.hits.length) {
      out.hits.set(body, found.hits)
      out.count += found.hits.length
    }

    if (out.count >= limit) {
      break
    }
  }

  return out
}

function touched(ctx: DeckText, matched: Matched): TextShape[] {
  const seen = new Map<string, TextShape>()
  for (const body of matched.hits.keys()) {
    ctx.touch(body.part)
    seen.set(`${body.slide}:${body.shape}`, { slide: body.slide, shape: body.shape })
  }

  return [...seen.values()]
}

function withRefused<T extends object>(result: T, refused: number): T & Refused {
  return refused ? { ...result, refused } : result
}

export function replaceText(ctx: DeckText, input: ReplaceTextOptions): ReplacedText {
  if (!isRecord(input)) {
    throw new DocsError('INVALID_OPTIONS', 'replaceText takes { find, replace, all, matchCase }')
  }

  const wanted = ranges.pattern(input)
  const replacement = ranges.plainText(input.replace, 'replace')
  const matched = match(scopedBodies(ctx, input), wanted)
  for (const hits of matched.hits.values()) {
    ranges.replaceHits(hits, replacement)
  }

  const shapes = touched(ctx, matched)

  return withRefused({ replaced: matched.count, shapes }, matched.refused)
}

function targetHits(ctx: DeckText, target: ShapeRef, search: Search): Matched {
  const wanted = ranges.pattern(search)
  const bodies = targetBodies(ctx, target)
  const matched = match(bodies, wanted)
  if (matched.count === 0 && matched.refused === 0) {
    const { slide, shape } = bodies[0]
    throw new DocsError(
      'TEXT_NOT_FOUND',
      `shape ${shape} on slide ${slide} has no text matching "${wanted.find}"`,
      { field: 'find', got: wanted.find }
    )
  }

  return matched
}

function isolated(ctx: DeckText, matched: Matched): Element[] {
  touched(ctx, matched)
  return [...matched.hits.values()].flatMap((hits) => ranges.isolateHits(hits))
}

export function formatText(ctx: DeckText, request: FormatRequest): FormattedText {
  const { target, search } = request
  const style = ranges.textStyle(request.style)
  const matched = targetHits(ctx, target, search)
  const runs = isolated(ctx, matched)
  ranges.styleRuns(runs, style)

  return withRefused({ formatted: matched.count, runs: runs.length }, matched.refused)
}

function checkUrl(url: unknown): string {
  if (typeof url !== 'string' || !URL_PATTERN.test(url)) {
    throw new DocsError(
      'INVALID_LINK',
      `a link must be an http, https or mailto URL, got ${JSON.stringify(url)}`,
      { field: 'url', expected: 'http://, https:// or mailto: URL', got: url }
    )
  }

  return url
}

function hyperlinkRel(ctx: DeckText, part: string, url: string): string {
  const rels = ctx.deck.pkg.rels(part)
  const existing = rels.all().find((r) => r.type === HYPERLINK && r.external && r.target === url)
  if (existing) {
    return existing.id
  }

  const taken = new Set(rels.all().map((r) => r.id))
  let n = taken.size + 1
  while (taken.has(`rId${n}`)) {
    n++
  }

  const id = `rId${n}`
  const el = xml.named(rels.doc, 'Relationship', xml.ns.rels)
  rels.root.appendChild(
    xml.setAttributes(el, { Id: id, Type: HYPERLINK, Target: url, TargetMode: 'External' })
  )
  rels.dirty = true

  return id
}

export function link(ctx: DeckText, request: LinkRequest): LinkedText {
  const href = checkUrl(request.url)
  const matched = targetHits(ctx, request.target, request.search)
  const runs = isolated(ctx, matched)
  const [body] = matched.hits.keys()
  if (body) {
    ranges.linkRuns(runs, hyperlinkRel(ctx, body.part, href))
  }

  return withRefused({ linked: matched.count, runs: runs.length }, matched.refused)
}
