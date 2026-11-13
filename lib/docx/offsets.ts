import * as xml from '../xml.ts'
import * as text from './text.ts'
import * as ranges from './ranges.ts'
import * as match from './match.ts'
import { ownFormat, places, segments } from './runs.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Element, Node } from '../xml.ts'
import type { RunFormat } from './text.ts'
import type { Styles } from './styles.ts'
import type { LinkTarget, RunPlace, Segment } from './runs.ts'
import type { Range } from './match.ts'

export interface InsertTextOptions {
  offset: number
  text: string
  format?: RunFormat
}

export interface TextRange {
  start: number
  end: number
}

export interface OffsetOptions {
  offset: number
}

export interface FormatRangeOptions extends TextRange {
  format: RunFormat
}

export interface TextEdit {
  length: number
  offset: number
}

export interface RangeFormatted {
  changed: boolean
}

export interface CaretFormat {
  bold: boolean
  italic: boolean
  underline: boolean
  size?: number
  color?: string
  font?: string
  link?: string
}

export interface Paragraph {
  p: Element
  block: number
}

export interface Context {
  styles: Styles
  target: LinkTarget
  touch: () => void
}

interface Caret {
  list: RunPlace[]
  host: Segment | null
  inside: boolean
  after: boolean
  other: Segment | null
}

interface Typing {
  rPr: Element | null
  link: Element | null
  escape: Element | null
}

interface Checked {
  field: string
  value: unknown
}

interface Insertion {
  caret: Caret
  offset: number
  value: string
}

const W = xml.ns.w
const ESCAPED = new Set(['hyperlink', 'fldSimple'])
const EMPTIED = new Set(['hyperlink', 'smartTag', 'customXml', 'dir', 'bdo'])
const TRACKED = new Set(['ins', 'del', 'moveFrom', 'moveTo'])
const PROPERTIES = new Set(['pPr', 'rPr', 'smartTagPr', 'customXmlPr', 'sdtPr', 'sdtEndPr'])
const MARK_REVISIONS = ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange']
const LINK_FORMAT = ['rStyle', 'color', 'u']

function option(opts: unknown, key: string): unknown {
  return isRecord(opts) ? opts[key] : undefined
}

function isHigh(code: number): boolean {
  return code >= 0xd800 && code <= 0xdbff
}

function isLow(code: number): boolean {
  return code >= 0xdc00 && code <= 0xdfff
}

function offsetOf(para: Paragraph, { field, value }: Checked): number {
  const content = text.paragraphText(para.p)
  const { length } = content
  if (typeof value !== 'number' || !Number.isInteger(value) || value < 0 || value > length) {
    throw new DocsError(
      'INVALID_OFFSET',
      `${field} must be a whole number from 0 to ${length}, the length of the text in block ${para.block}, got ${JSON.stringify(value)}`,
      { field, expected: `0 to ${length}`, got: value, block: para.block, length }
    )
  }

  if (isHigh(content.charCodeAt(value - 1)) && isLow(content.charCodeAt(value))) {
    throw new DocsError(
      'INVALID_OFFSET',
      `${field} ${value} falls between the two halves of a character in block ${para.block}; use ${value - 1} or ${value + 1}`,
      { field, got: value, block: para.block, candidates: [value - 1, value + 1] }
    )
  }

  return value
}

function rangeOf(para: Paragraph, opts: unknown): Range {
  const start = offsetOf(para, { field: 'start', value: option(opts, 'start') })
  const end = offsetOf(para, { field: 'end', value: option(opts, 'end') })
  if (end < start) {
    throw new DocsError(
      'INVALID_OFFSET',
      `end ${end} is before start ${start} in block ${para.block}`,
      { field: 'end', expected: `${start} or more`, got: end, block: para.block }
    )
  }

  return [start, end]
}

function insertedText(value: unknown): string {
  if (typeof value !== 'string') {
    throw new DocsError('INVALID_TEXT', 'text must be the text to insert', {
      field: 'text',
      expected: 'string',
      got: value
    })
  }

  return xml.clean(value.replace(/\r\n?|\v/g, '\n'))
}

function formatOption(opts: unknown): RunFormat | null {
  const input = option(opts, 'format')
  return input === undefined ? null : text.validateFormat(input, text.RUN_KEYS)
}

function caretAt(p: Element, offset: number): Caret {
  const list = places(p)
  const segs = segments(list)
  const inside = segs.find((seg) => seg.start < offset && offset < seg.end)
  if (inside) {
    return { list, host: inside, inside: true, after: false, other: null }
  }

  const before = segs.find((seg) => seg.end === offset)
  const next = segs.find((seg) => seg.start === offset) ?? null
  if (before) {
    return { list, host: before, inside: false, after: true, other: next }
  }

  return { list, host: next, inside: false, after: false, other: null }
}

function inField(caret: Caret): boolean {
  if (!caret.host?.place.field) {
    return false
  }

  return caret.inside || !!caret.other?.place.field
}

function refuseField(para: Paragraph, offset: number): DocsError {
  return new DocsError(
    'FIELD_TEXT',
    `offset ${offset} in block ${para.block} is inside a field result, which Word recalculates; edit the field instead`,
    { field: 'offset', got: offset, block: para.block }
  )
}

function refuseTracked(para: Paragraph, offset: number): DocsError {
  return new DocsError(
    'TRACKED_CHANGES',
    `the text at offset ${offset} in block ${para.block} has tracked changes; accept or reject them in Word first`,
    { field: 'offset', got: offset, block: para.block }
  )
}

function within(node: Node | null, container: Element): boolean {
  for (let el = node; el; el = el.parentNode) {
    if (el === container) {
      return true
    }
  }

  return false
}

function fieldEdge(caret: Caret, host: Segment): Element {
  const { list } = caret
  const step = caret.after ? 1 : -1
  let at = list.indexOf(host.place)
  while (list[at + step]?.field) {
    at += step
  }

  return list[at].run
}

function escapeNode(p: Element, caret: Caret, host: Segment): Element | null {
  const other = caret.other
  if (caret.inside) {
    return null
  }

  let node = host.place.field ? fieldEdge(caret, host) : host.place.run
  for (let el = node.parentNode; xml.isElement(el) && el !== p; el = el.parentNode) {
    const escapes = xml.isElement(el, W) && ESCAPED.has(el.localName ?? '')
    if (escapes && !(other && within(other.node, el))) {
      node = el
    }
  }

  return node === host.place.run ? null : node
}

function withoutRevisions(rPr: Element): Element {
  for (const local of MARK_REVISIONS) {
    for (const el of xml.children(rPr, W, local)) {
      xml.remove(el)
    }
  }

  return rPr
}

function withoutLinkFormat(rPr: Element): Element {
  for (const local of LINK_FORMAT) {
    for (const el of xml.children(rPr, W, local)) {
      xml.remove(el)
    }
  }

  return rPr
}

function markRunProperties(p: Element): Element | null {
  const pPr = xml.child(p, W, 'pPr')
  const mark = pPr ? xml.child(pPr, W, 'rPr') : null
  return mark ? withoutRevisions(xml.clone(mark)) : null
}

function typing(p: Element, caret: Caret): Typing {
  const host = caret.host
  if (!host) {
    return { rPr: markRunProperties(p), link: null, escape: null }
  }

  const own = xml.child(host.place.run, W, 'rPr')
  const rPr = own ? xml.clone(own) : null
  const escape = escapeNode(p, caret, host)
  const link = host.place.link
  if (escape && link && within(link, escape)) {
    return { rPr: rPr && withoutLinkFormat(rPr), link: null, escape }
  }

  return { rPr, link, escape }
}

function styleIdOf(p: Element): string | null {
  const pPr = xml.child(p, W, 'pPr')
  return text.wGet(pPr ? xml.child(pPr, W, 'pStyle') : null, 'val') || null
}

function effective(p: Element, rPr: Element | null, styles: Styles): RunFormat {
  const characterStyle = text.wGet(rPr ? xml.child(rPr, W, 'rStyle') : null, 'val') || null
  const layers = [...styles.runLayers(styleIdOf(p), characterStyle), rPr].filter(
    (el) => el !== null
  )

  return layers.reduce<RunFormat>((out, el) => ({ ...out, ...ownFormat(el) }), {})
}

function caretFormat(found: RunFormat, link: string | undefined): CaretFormat {
  const out: CaretFormat = {
    bold: found.bold === true,
    italic: found.italic === true,
    underline: found.underline === true
  }
  if (found.size !== undefined) {
    out.size = found.size
  }

  if (found.color && found.color !== 'auto') {
    out.color = found.color.toUpperCase()
  }

  if (found.font) {
    out.font = found.font
  }

  if (link) {
    out.link = link
  }

  return out
}

function sameAs(current: CaretFormat, format: RunFormat): boolean {
  return Object.entries(format).every(([key, value]) => {
    const now: unknown = Reflect.get(current, key)
    if (key === 'color') {
      return String(now) === String(value).replace('#', '').toUpperCase()
    }

    return now === value
  })
}

function nextElement(node: Node): Element | null {
  for (let n = node.nextSibling; n; n = n.nextSibling) {
    if (xml.isElement(n)) {
      return n
    }
  }

  return null
}

function firstContent(el: Element): Element | null {
  return xml.children(el).find((child) => !PROPERTIES.has(child.localName ?? '')) ?? null
}

function insertAfter(node: Element, added: Element): void {
  node.parentNode?.insertBefore(added, node.nextSibling)
}

function splitPoint({ caret, offset }: Insertion, host: Segment): Element | null {
  if (caret.inside) {
    return ranges.splitText(host.node, offset - host.start)
  }

  return caret.after ? nextElement(host.node) : host.node
}

function runWith(p: Element, rPr: Element | null, value: string): Element {
  const doc = xml.owner(p)
  const run = text.wEl(doc, 'r')
  if (rPr && xml.children(rPr).length) {
    run.appendChild(rPr)
  }

  for (const el of ranges.pieces(doc, value)) {
    run.appendChild(el)
  }

  return run
}

function placeIn(run: Element, before: Element | null, added: Element): void {
  if (before && firstContent(run) === before) {
    run.parentNode?.insertBefore(added, run)
    return
  }

  if (before) {
    ranges.splitRunBefore(before)
  }

  insertAfter(run, added)
}

function typeInto(insertion: Insertion, host: Segment): void {
  const { node } = host
  if (node.localName === 't' && !/[\n\t]/.test(insertion.value)) {
    const current = xml.text(node)
    const at = insertion.offset - host.start
    ranges.setT(node, current.slice(0, at) + insertion.value + current.slice(at))
    return
  }

  const before = splitPoint(insertion, host)
  for (const el of ranges.pieces(xml.owner(node), insertion.value)) {
    node.parentNode?.insertBefore(el, before)
  }
}

function placeEscaped(insertion: Insertion, escape: Element, added: Element): void {
  if (insertion.caret.after) {
    insertAfter(escape, added)
  } else {
    escape.parentNode?.insertBefore(added, escape)
  }
}

function insertAt(para: Paragraph, insertion: Insertion, format: RunFormat | null): void {
  const { p } = para
  const { caret, value } = insertion
  const host = caret.host
  const typed = typing(p, caret)
  const rPr = typed.rPr ?? text.wEl(xml.owner(p), 'rPr')
  if (format) {
    text.applyRunFormat(rPr, format)
  }

  if (!host) {
    p.appendChild(runWith(p, rPr, value))
  } else if (typed.escape) {
    placeEscaped(insertion, typed.escape, runWith(p, rPr, value))
  } else if (format) {
    placeIn(host.place.run, splitPoint(insertion, host), runWith(p, rPr, value))
  } else {
    typeInto(insertion, host)
  }
}

export function insertText(para: Paragraph, opts: unknown, context: Context): TextEdit {
  const offset = offsetOf(para, { field: 'offset', value: option(opts, 'offset') })
  const value = insertedText(option(opts, 'text'))
  const requested = formatOption(opts)
  if (value) {
    const caret = caretAt(para.p, offset)
    if (caret.host?.place.tracked) {
      throw refuseTracked(para, offset)
    }

    if (inField(caret)) {
      throw refuseField(para, offset)
    }

    const current = caretFormat(
      effective(para.p, typing(para.p, caret).rPr, context.styles),
      undefined
    )
    const format = requested && !sameAs(current, requested) ? requested : null
    insertAt(para, { caret, offset, value }, format)
    context.touch()
  }

  return { length: text.paragraphText(para.p).length, offset: offset + value.length }
}

function touchedContainers(p: Element, range: Range): Element[] {
  const out = new Set<Element>()
  for (const seg of match.overlapping(segments(places(p)), range)) {
    for (let el = seg.place.run.parentNode; el && el !== p; el = el.parentNode) {
      if (xml.isElement(el, W) && EMPTIED.has(el.localName ?? '')) {
        out.add(el)
      }
    }
  }

  return [...out]
}

export function deleteText(para: Paragraph, opts: unknown, context: Context): TextEdit {
  const range = rangeOf(para, opts)
  if (range[0] < range[1]) {
    match.checkAll([{ index: para.block, p: para.p, ranges: [range] }], 'delete')
    const containers = touchedContainers(para.p, range)
    ranges.replaceIn(para.p, [range], '')
    for (const el of containers) {
      if (el.parentNode && !firstContent(el)) {
        xml.remove(el)
      }
    }

    context.touch()
  }

  return { length: text.paragraphText(para.p).length, offset: range[0] }
}

export function formatRange(para: Paragraph, opts: unknown, context: Context): RangeFormatted {
  const range = rangeOf(para, opts)
  const format = text.validateFormat(option(opts, 'format'), text.RUN_KEYS)
  if (range[0] === range[1]) {
    return { changed: false }
  }

  match.checkAll([{ index: para.block, p: para.p, ranges: [range] }], 'format')
  ranges.formatIn(para.p, [range], format)
  context.touch()

  return { changed: true }
}

export function runFormat(para: Paragraph, opts: unknown, context: Context): CaretFormat {
  const offset = offsetOf(para, { field: 'offset', value: option(opts, 'offset') })
  const typed = typing(para.p, caretAt(para.p, offset))
  const link = typed.link ? context.target(typed.link) : undefined
  return caretFormat(effective(para.p, typed.rPr, context.styles), link)
}

function refuseMark(para: Paragraph): void {
  const pPr = xml.child(para.p, W, 'pPr')
  const mark = pPr ? xml.child(pPr, W, 'rPr') : null
  const tracked =
    (!!pPr && !!xml.child(pPr, W, 'pPrChange')) ||
    (!!mark && MARK_REVISIONS.some((local) => xml.child(mark, W, local)))
  if (tracked) {
    throw new DocsError(
      'TRACKED_CHANGES',
      `the paragraph mark of block ${para.block} has tracked changes; accept or reject them in Word first`,
      { block: para.block }
    )
  }
}

function fieldStart(p: Element, run: Element): Element {
  const open: Element[] = []
  for (const place of places(p)) {
    if (place.run === run) {
      break
    }

    for (const mark of xml.children(place.run, W, 'fldChar')) {
      const type = mark.getAttributeNS(W, 'fldCharType')
      if (type === 'begin') {
        open.push(place.run)
      } else if (type === 'end') {
        open.pop()
      }
    }
  }

  return open[0] ?? run
}

function refuseSplit(para: Paragraph, container: Element): void {
  const name = container.localName ?? ''
  if (TRACKED.has(name)) {
    throw new DocsError(
      'TRACKED_CHANGES',
      `the split point in block ${para.block} is inside a tracked change; accept or reject it in Word first`,
      { block: para.block }
    )
  }

  if (name === 'fldSimple') {
    throw new DocsError('FIELD_TEXT', `the split point in block ${para.block} is inside a field`, {
      block: para.block
    })
  }

  if (name === 'sdt' || name === 'sdtContent') {
    throw new DocsError(
      'RANGE_BOUNDARY',
      `the split point in block ${para.block} is inside a content control`,
      { block: para.block }
    )
  }
}

function splitContainer(para: Paragraph, container: Element, at: Element): Element {
  refuseSplit(para, container)
  const doc = xml.owner(container)
  const right = doc.createElementNS(container.namespaceURI, container.tagName)
  for (const attr of Array.from(container.attributes)) {
    right.setAttributeNS(attr.namespaceURI, attr.name, attr.value)
  }

  for (const prop of xml.children(container).filter((el) => PROPERTIES.has(el.localName ?? ''))) {
    right.appendChild(xml.clone(prop))
  }

  for (const n of ranges.siblingsFrom(at, null)) {
    right.appendChild(n)
  }

  insertAfter(container, right)

  return right
}

function lift(para: Paragraph, node: Element): Element {
  let at = node
  let parent = at.parentNode
  while (xml.isElement(parent) && parent !== para.p) {
    at = firstContent(parent) === at ? parent : splitContainer(para, parent, at)
    parent = at.parentNode
  }

  return at
}

function cutFrom(para: Paragraph, offset: number): Element | null {
  if (offset === 0) {
    return firstContent(para.p)
  }

  ranges.splitAt(para.p, offset)
  const seg = segments(places(para.p)).find((item) => item.start === offset)

  return seg ? lift(para, fieldStart(para.p, seg.place.run)) : null
}

function applyNext(next: Element, current: string | null, styles: Styles): void {
  const id = styles.nextOf(current)
  const now = current ?? styles.defaultParagraph()?.id ?? null
  if (!id || id === now) {
    return
  }

  const pPr = text.ensure(next, 'pPr')
  text.wAttr(text.ensure(pPr, 'pStyle', text.PARAGRAPH_ORDER), 'val', id)
}

export function splitParagraph(para: Paragraph, opts: unknown, context: Context): Element {
  const offset = offsetOf(para, { field: 'offset', value: option(opts, 'offset') })
  refuseMark(para)
  const { p } = para
  const length = text.paragraphText(p).length
  if (inField(caretAt(p, offset))) {
    throw refuseField(para, offset)
  }

  const cut = cutFrom(para, offset)
  const next = text.wEl(xml.owner(p), 'p')
  const pPr = xml.child(p, W, 'pPr')
  if (pPr) {
    next.appendChild(xml.clone(pPr))
    for (const sect of xml.children(pPr, W, 'sectPr')) {
      xml.remove(sect)
    }
  }

  for (const n of cut ? ranges.siblingsFrom(cut, null) : []) {
    next.appendChild(n)
  }

  insertAfter(p, next)
  if (offset === length) {
    applyNext(next, styleIdOf(p), context.styles)
  }

  context.touch()

  return next
}

export function mergeParagraphs(para: Paragraph, next: Element, context: Context): number {
  refuseMark(para)
  refuseMark({ p: next, block: para.block + 1 })
  const { p } = para
  const offset = text.paragraphText(p).length
  const nextPr = xml.child(next, W, 'pPr')
  const sect = nextPr ? xml.child(nextPr, W, 'sectPr') : null
  for (const n of xml.children(next).filter((el) => el !== nextPr)) {
    p.appendChild(n)
  }

  if (sect) {
    const pPr = text.ensure(p, 'pPr')
    for (const old of xml.children(pPr, W, 'sectPr')) {
      xml.remove(old)
    }

    xml.insertOrdered(pPr, sect, text.PARAGRAPH_ORDER)
  }

  xml.remove(next)
  context.touch()

  return offset
}
