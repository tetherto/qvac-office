import * as xml from '../xml.ts'
import * as text from './text.ts'
import * as match from './match.ts'
import { places, segments, type Segment } from './runs.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Document as XmlDocument, Element, Node } from '../xml.ts'
import type { RunFormat } from './text.ts'
import type { BlockRef } from './ids.ts'
import type { Range, Target, TextQuery } from './match.ts'

export interface ReplaceOptions extends TextQuery {
  replace: string
  in?: BlockRef
}

export interface ReplaceResult {
  replaced: number
  blocks: number[]
}

export interface FormatTextResult {
  formatted: number
}

export interface LinkResult {
  linked: number
}

export interface LinkPlan {
  rel: () => string
  style: string | null
}

interface Located {
  p: Element
  range: Range
}

const W = xml.ns.w
const R = xml.ns.r
const XML_NS = 'http://www.w3.org/XML/1998/namespace'
const LINK_COLOR = '0563C1'

function setT(t: Element, value: string): void {
  xml.setText(t, value)
  if (/^\s|\s$/.test(value)) {
    t.setAttributeNS(XML_NS, 'xml:space', 'preserve')
  }
}

function tEl(doc: XmlDocument, value: string): Element {
  const t = text.wEl(doc, 't')
  setT(t, value)
  return t
}

function piece(doc: XmlDocument, value: string): Element {
  if (value === '\n') {
    return text.wEl(doc, 'br')
  }

  return value === '\t' ? text.wEl(doc, 'tab') : tEl(doc, value)
}

function pieces(doc: XmlDocument, value: string): Element[] {
  return value
    .replace(/\r\n?/g, '\n')
    .split(/(\n|\t)/)
    .filter(Boolean)
    .map((part) => piece(doc, part))
}

function insertAfter(node: Element, added: Element[]): void {
  let before = node.nextSibling
  for (const el of added) {
    node.parentNode?.insertBefore(el, before)
    before = el.nextSibling
  }
}

function writeStart(seg: Segment, [a, b]: Range, replacement: string): void {
  const { node } = seg
  const doc = xml.owner(node)
  if (node.localName !== 't') {
    for (const el of pieces(doc, replacement)) {
      node.parentNode?.insertBefore(el, node)
    }

    xml.remove(node)

    return
  }

  const current = xml.text(node)
  if (!/[\n\r\t]/.test(replacement)) {
    setT(node, current.slice(0, a) + replacement + current.slice(b))
    return
  }

  const after = current.slice(b)
  setT(node, current.slice(0, a))
  insertAfter(node, [...pieces(doc, replacement), ...(after ? [tEl(doc, after)] : [])])
}

function cut(seg: Segment, [a, b]: Range): void {
  if (seg.node.localName !== 't') {
    xml.remove(seg.node)
    return
  }

  const current = xml.text(seg.node)
  setT(seg.node, current.slice(0, a) + current.slice(b))
}

function tidy(runs: Element[]): void {
  for (const run of runs) {
    for (const t of xml.children(run, W, 't').filter((el) => xml.text(el) === '')) {
      xml.remove(t)
    }

    if (xml.children(run).every((el) => el.localName === 'rPr')) {
      xml.remove(run)
    }
  }
}

function local(seg: Segment, [start, end]: Range): Range {
  return [Math.max(start, seg.start) - seg.start, Math.min(end, seg.end) - seg.start]
}

function replaceIn(p: Element, ranges: Range[], replacement: string): void {
  const segs = segments(places(p))
  for (const range of [...ranges].reverse()) {
    const touched = match.overlapping(segs, range)
    writeStart(touched[0], local(touched[0], range), replacement)
    for (const seg of touched.slice(1)) {
      cut(seg, local(seg, range))
    }

    tidy([...new Set(touched.map((seg) => seg.place.run))])
  }
}

function shell(run: Element): Element {
  const copy = text.wEl(xml.owner(run), 'r')
  for (const attr of Array.from(run.attributes)) {
    copy.setAttributeNS(attr.namespaceURI, attr.name, attr.value)
  }

  const rPr = xml.child(run, W, 'rPr')
  if (rPr) {
    copy.appendChild(xml.clone(rPr))
  }

  return copy
}

function siblingsFrom(first: Node, last: Node | null): Node[] {
  const out: Node[] = []
  for (let n: Node | null = first; n; n = n.nextSibling) {
    out.push(n)
    if (n === last) {
      break
    }
  }

  return out
}

function splitRunBefore(node: Element): void {
  const run = node.parentNode
  if (!xml.isElement(run) || xml.children(run).find((el) => el.localName !== 'rPr') === node) {
    return
  }

  const right = shell(run)
  for (const n of siblingsFrom(node, null)) {
    right.appendChild(n)
  }

  run.parentNode?.insertBefore(right, run.nextSibling)
}

function splitText(t: Element, at: number): Element {
  const current = xml.text(t)
  const tail = tEl(xml.owner(t), current.slice(at))
  setT(t, current.slice(0, at))
  t.parentNode?.insertBefore(tail, t.nextSibling)
  return tail
}

function splitAt(p: Element, at: number): void {
  const segs = segments(places(p))
  const inside = segs.find((seg) => seg.start < at && at < seg.end)
  if (inside) {
    splitRunBefore(splitText(inside.node, at - inside.start))
    return
  }

  const next = segs.find((seg) => seg.start === at)
  if (next) {
    splitRunBefore(next.node)
  }
}

function isolate(p: Element, [start, end]: Range): Element[] {
  splitAt(p, start)
  splitAt(p, end)
  const inside = segments(places(p)).filter((seg) => seg.start >= start && seg.end <= end)
  return [...new Set(inside.map((seg) => seg.place.run))]
}

function formatIn(p: Element, ranges: Range[], format: RunFormat): void {
  for (const run of ranges.flatMap((range) => isolate(p, range))) {
    text.applyRunFormat(text.ensure(run, 'rPr'), format)
  }
}

function linkFormat(run: Element, style: string | null): void {
  const rPr = text.ensure(run, 'rPr')
  if (style) {
    text.wAttr(text.ensure(rPr, 'rStyle', text.RUN_ORDER), 'val', style)
  } else {
    text.applyRunFormat(rPr, { color: LINK_COLOR, underline: true })
  }
}

function wrapLink({ p, range }: Located, plan: LinkPlan): void {
  const runs = isolate(p, range)
  const link = text.wEl(xml.owner(p), 'hyperlink', { history: 1 })
  xml.setNsAttribute(link, { namespace: R, local: 'id' }, plan.rel())
  runs[0].parentNode?.insertBefore(link, runs[0])
  for (const n of siblingsFrom(runs[0], runs[runs.length - 1])) {
    link.appendChild(n)
  }

  for (const run of runs) {
    linkFormat(run, plan.style)
  }
}

export function replaceText(targets: Target[], opts: ReplaceOptions): ReplaceResult {
  const query = match.validateQuery(opts, 'options')
  const replacement: unknown = isRecord(opts) ? opts.replace : undefined
  if (typeof replacement !== 'string') {
    throw new DocsError('INVALID_OPTIONS', 'options.replace must be the replacement text', {
      field: 'options.replace',
      expected: 'string',
      got: replacement
    })
  }

  const planned = match.plan(targets, query)
  match.checkAll(planned, 'replace')
  for (const { p, ranges } of planned) {
    replaceIn(p, ranges, replacement)
  }

  return { replaced: match.count(planned), blocks: [...new Set(planned.map((item) => item.index))] }
}

function found(targets: Target[], planned: match.Planned[]): match.Planned[] {
  if (planned.length) {
    return planned
  }

  const [index, block] = targets[0]

  throw new DocsError('TEXT_NOT_FOUND', `block ${index} has no text matching that search`, {
    field: 'find',
    block: index,
    text: text.blockText(block).slice(0, 200)
  })
}

export function formatText(targets: Target[], query: unknown, input: unknown): FormatTextResult {
  const planned = found(targets, match.plan(targets, match.validateQuery(query)))
  const format = text.validateFormat(input, text.RUN_KEYS)
  match.checkAll(planned, 'format')
  for (const { p, ranges } of planned) {
    formatIn(p, ranges, format)
  }

  return { formatted: match.count(planned) }
}

export function linkText(targets: Target[], query: unknown, plan: LinkPlan): LinkResult {
  const [first] = found(targets, match.plan(targets, { ...match.validateQuery(query), all: false }))
  match.checkAll([first], 'link')
  wrapLink({ p: first.p, range: first.ranges[0] }, plan)

  return { linked: 1 }
}
