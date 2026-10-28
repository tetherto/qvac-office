import * as xml from '../xml.ts'
import * as text from './text.ts'
import { DocsError } from '../errors.ts'
import { W } from './host.ts'
import type { Element, Node } from '../xml.ts'

export interface Span {
  first: Element
  last: Element
}

interface Piece {
  run: Element
  start: number
  length: number
}

const HIDDEN = new Set(['del', 'moveFrom', 'txbxContent', 'pict', 'drawing', 'object'])
const XML_NS = 'http://www.w3.org/XML/1998/namespace'

function visible(r: Element, p: Element): boolean {
  for (let node = r.parentNode; node && node !== p; node = node.parentNode) {
    if (xml.isElement(node, W) && HIDDEN.has(node.localName ?? '')) {
      return false
    }
  }

  return true
}

function childLength(el: Element): number {
  const name = el.localName
  if (name === 't') {
    return xml.text(el).length
  }

  if (name === 'br') {
    return text.wGet(el, 'type') === 'page' ? 0 : 1
  }

  return name === 'tab' || name === 'cr' || name === 'noBreakHyphen' ? 1 : 0
}

function runLength(r: Element): number {
  return xml.children(r, W).reduce((sum, el) => sum + childLength(el), 0)
}

function pieces(p: Element): Piece[] {
  const out: Piece[] = []
  let at = 0
  for (const run of xml.descendants(p, W, 'r').filter((r) => visible(r, p))) {
    const length = runLength(run)
    out.push({ run, start: at, length })
    at += length
  }

  return out
}

function textNode(doc: ReturnType<typeof xml.owner>, value: string): Element {
  const t = text.wEl(doc, 't')
  t.setAttributeNS(XML_NS, 'xml:space', 'preserve')
  t.appendChild(doc.createTextNode(value))
  return t
}

function moveAfter(r: Element, offset: number): Element {
  const doc = xml.owner(r)
  const tail = text.wEl(doc, 'r')
  const rPr = xml.child(r, W, 'rPr')
  if (rPr) {
    tail.appendChild(xml.clone(rPr))
  }

  let at = 0
  const moved: Node[] = []
  for (const el of xml.children(r, W)) {
    if (el.localName === 'rPr') {
      continue
    }

    const length = childLength(el)
    if (at >= offset) {
      moved.push(el)
    } else if (el.localName === 't' && at + length > offset) {
      const value = xml.text(el)
      xml.remove(el)
      r.appendChild(textNode(doc, value.slice(0, offset - at)))
      tail.appendChild(textNode(doc, value.slice(offset - at)))
    }

    at += length
  }

  for (const node of moved) {
    tail.appendChild(node)
  }

  r.parentNode?.insertBefore(tail, r.nextSibling)

  return tail
}

function splitAt(p: Element, offset: number): void {
  const piece = pieces(p).find((x) => offset > x.start && offset < x.start + x.length)
  if (piece) {
    moveAfter(piece.run, offset - piece.start)
  }
}

export function locate(p: Element, find: string, where: string): Span {
  const all = pieces(p)
  const whole = all.map((x) => text.paragraphText(holder(x.run))).join('')
  const start = whole.indexOf(find)
  if (start === -1 || !find) {
    throw new DocsError(
      'TEXT_NOT_FOUND',
      `"${find}" is not in ${where}; its text is "${whole.slice(0, 200)}"`,
      { field: 'find', got: find, text: whole.slice(0, 200) }
    )
  }

  const end = start + find.length
  splitAt(p, start)
  splitAt(p, end)
  const inside = pieces(p).filter((x) => x.length > 0 && x.start >= start && x.start < end)

  return { first: inside[0].run, last: inside[inside.length - 1].run }
}

function holder(r: Element): Element {
  const p = text.wEl(xml.owner(r), 'p')
  p.appendChild(xml.clone(r))
  return p
}

export function contentStart(p: Element): Node | null {
  const pPr = xml.child(p, W, 'pPr')
  return pPr ? pPr.nextSibling : p.firstChild
}
