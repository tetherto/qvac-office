import * as xml from '../xml.ts'
import * as text from './text.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Element } from '../xml.ts'

export interface Search {
  find: string
  all?: boolean
  matchCase?: boolean
}

export interface Pattern {
  find: string
  all: boolean
  regex: RegExp
}

interface Piece {
  el: Element
  start: number
  end: number
  editable: boolean
}

export interface Hit {
  pieces: Piece[]
  start: number
  end: number
}

export interface Found {
  hits: Hit[]
  refused: number
}

const { a: A, r: R } = xml.ns

export function escapeRegex(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function flag(value: unknown, name: string): boolean {
  if (value !== undefined && typeof value !== 'boolean') {
    throw new DocsError('INVALID_OPTIONS', `${name} must be true or false`, {
      field: name,
      expected: 'boolean',
      got: value
    })
  }

  return value === true
}

export function plainText(value: unknown, name: string): string {
  if (typeof value !== 'string' || /[\r\n\v]/.test(value)) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `${name} must be text on one line; use setText to change whole paragraphs`,
      { field: name, expected: 'string without line breaks', got: value }
    )
  }

  return value
}

export function pattern(search: unknown): Pattern {
  if (!isRecord(search)) {
    throw new DocsError('INVALID_OPTIONS', 'the search must be { find, all, matchCase }')
  }

  const find = plainText(search.find, 'find')
  if (find === '') {
    throw new DocsError('INVALID_OPTIONS', 'find must not be empty', { field: 'find', got: find })
  }

  const matchCase = flag(search.matchCase, 'matchCase')
  const all = flag(search.all, 'all')

  return { find, all, regex: new RegExp(escapeRegex(find), matchCase ? 'gu' : 'giu') }
}

function pieces(p: Element): Piece[] {
  const out: Piece[] = []
  let at = 0
  for (const node of xml.children(p, A)) {
    const local = node.localName
    if (local !== 'r' && local !== 'fld' && local !== 'br') {
      continue
    }

    const length = local === 'br' ? 1 : xml.text(xml.child(node, A, 't')).length
    out.push({ el: node, start: at, end: at + length, editable: local === 'r' })
    at += length
  }

  return out
}

function covered(hit: Hit): Piece[] {
  return hit.pieces.filter((piece) => piece.start < hit.end && piece.end > hit.start)
}

function pieceText(piece: Piece): string {
  return piece.el.localName === 'br' ? '\n' : xml.text(xml.child(piece.el, A, 't'))
}

function paragraphHits(p: Element, wanted: Pattern): Hit[] {
  const list = pieces(p)
  const value = list.map(pieceText).join('')
  return [...value.matchAll(wanted.regex)].map((match) => ({
    pieces: list,
    start: match.index,
    end: match.index + match[0].length
  }))
}

export function search(txBody: Element, wanted: Pattern, limit: number): Found {
  const out: Found = { hits: [], refused: 0 }
  const all = xml.children(txBody, A, 'p').flatMap((p) => paragraphHits(p, wanted))
  for (const hit of all) {
    if (out.hits.length >= limit) {
      break
    }

    if (covered(hit).every((piece) => piece.editable)) {
      out.hits.push(hit)
    } else {
      out.refused++
    }
  }

  return out
}

function textOf(run: Element): string {
  return xml.text(xml.child(run, A, 't'))
}

function setRunText(run: Element, value: string): void {
  const t = xml.child(run, A, 't') ?? xml.append(run, xml.builder(xml.owner(run), A)('t'))
  xml.setText(t, value)
}

function replaceHit(hit: Hit, replacement: string): void {
  const list = covered(hit)
  const first = list[0]
  const last = list[list.length - 1]
  const tail = textOf(last.el).slice(hit.end - last.start)
  const head = textOf(first.el).slice(0, hit.start - first.start)
  setRunText(first.el, head + replacement + (first === last ? tail : ''))
  for (const piece of list.slice(1, -1)) {
    xml.remove(piece.el)
  }

  if (first !== last) {
    setRunText(last.el, tail)
    if (tail === '') {
      xml.remove(last.el)
    }
  }

  if (textOf(first.el) === '') {
    xml.remove(first.el)
  }
}

export function replaceHits(hits: Hit[], replacement: string): void {
  for (const hit of [...hits].reverse()) {
    replaceHit(hit, xml.clean(replacement))
  }
}

function splitRun(run: Element, at: number): Element {
  const value = textOf(run)
  const right = xml.clone(run)
  setRunText(run, value.slice(0, at))
  setRunText(right, value.slice(at))
  run.parentNode?.insertBefore(right, run.nextSibling)
  return right
}

function isolate(hit: Hit): Element[] {
  const list = covered(hit)
  const first = list[0]
  const last = list[list.length - 1]
  if (hit.end < last.end) {
    splitRun(last.el, hit.end - last.start)
  }

  const head = hit.start > first.start ? splitRun(first.el, hit.start - first.start) : first.el

  return [head, ...list.slice(1).map((piece) => piece.el)]
}

export function isolateHits(hits: Hit[]): Element[] {
  return [...hits].reverse().flatMap(isolate)
}

function runProperties(run: Element): Element {
  const existing = xml.child(run, A, 'rPr')
  if (existing) {
    return existing
  }

  const rPr = xml.builder(xml.owner(run), A)('rPr')
  run.insertBefore(rPr, run.firstChild)

  return rPr
}

export function textStyle(input: unknown): text.Style {
  const style = text.validateStyle(input)
  if (style.align !== undefined) {
    throw new DocsError('INVALID_STYLE', 'align applies to whole paragraphs; use setStyle', {
      field: 'align'
    })
  }

  return style
}

export function styleRuns(runs: Element[], style: text.Style): void {
  for (const run of runs) {
    text.applyRunStyle(runProperties(run), style)
  }
}

export function linkRuns(runs: Element[], rId: string): void {
  for (const run of runs) {
    const rPr = runProperties(run)
    for (const old of xml.children(rPr, A, 'hlinkClick')) {
      xml.remove(old)
    }

    const hlink = xml.builder(xml.owner(run), A)('hlinkClick')
    xml.setNsAttribute(hlink, { namespace: R, local: 'id' }, rId)
    xml.insertOrdered(rPr, hlink, text.RUN_PROPERTY_ORDER)
  }
}
