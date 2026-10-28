import { DOMParser, XMLSerializer, Element, type Document, type Node } from '@xmldom/xmldom'
import { DocsError } from './errors.ts'
import { errorMessage } from './guards.ts'
import * as utf8 from './utf8.ts'

export type { Document, Element, Node }

export const ns = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  x: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  rels: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types'
} as const

export type Attributes = Record<string, string | number | undefined | null>

export type Build = (local: string, attributes?: Attributes) => Element

const MAX_DEPTH = 256
const NOT_XML =
  /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const UNPREFIXED: readonly string[] = [ns.rels, ns.ct, ns.x]
const PREFIXES = new Map<string, string>(
  Object.entries(ns)
    .filter(([, uri]) => !UNPREFIXED.includes(uri))
    .map(([prefix, uri]) => [uri, prefix])
)

const SKIPPED: readonly [string, string][] = [
  ['<!--', '-->'],
  ['<![CDATA[', ']]>'],
  ['<?', '?>'],
  ['<!', '>']
]

function skipped(text: string, at: number): number | null {
  for (const [open, close] of SKIPPED) {
    if (text.startsWith(open, at)) {
      const end = text.indexOf(close, at + open.length)
      return end === -1 ? -1 : end + close.length
    }
  }

  return null
}

function tagEnd(text: string, at: number): number {
  let quote = ''
  for (let i = at; i < text.length; i++) {
    const c = text[i]
    if (quote) {
      quote = c === quote ? '' : quote
    } else if (c === '"' || c === "'") {
      quote = c
    } else if (c === '>') {
      return i
    }
  }

  return -1
}

function checkDepth(text: string, name: string): void {
  let depth = 0
  for (let i = text.indexOf('<'); i !== -1; i = text.indexOf('<', i)) {
    const after = skipped(text, i)
    if (after === -1) {
      return
    }

    if (after !== null) {
      i = after
      continue
    }

    const end = tagEnd(text, i + 1)
    if (end === -1) {
      return
    }

    if (text[i + 1] === '/') {
      depth--
    } else if (text[end - 1] !== '/') {
      depth++
    }

    if (depth > MAX_DEPTH) {
      throw new DocsError('UNSAFE_XML', `${name} nests elements more than ${MAX_DEPTH} levels deep`)
    }

    i = end + 1
  }
}

function decodeUtf16(bytes: Uint8Array, endian: 'le' | 'be'): string {
  const littleEndian = endian === 'le'
  let out = ''
  for (let i = 2; i + 1 < bytes.length; i += 2) {
    const code = littleEndian ? bytes[i] | (bytes[i + 1] << 8) : (bytes[i] << 8) | bytes[i + 1]
    out += String.fromCharCode(code)
  }

  return out
}

export function decode(bytes: Uint8Array): string {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return utf8.decode(bytes.subarray(3))
  }

  if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    return decodeUtf16(bytes, 'le')
  }

  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    return decodeUtf16(bytes, 'be')
  }

  return utf8.decode(bytes)
}

function causeOf(err: unknown): unknown {
  return err instanceof Error ? err.cause : undefined
}

export function parse(input: Uint8Array | string, name = 'part'): Document {
  const text = (typeof input === 'string' ? input : decode(input)).replace(/^﻿/, '')
  if (/<!DOCTYPE/i.test(text)) {
    throw new DocsError('UNSAFE_XML', `${name} declares a DOCTYPE, which Office files never do`)
  }

  checkDepth(text, name)
  const invalid = (message: string): DocsError =>
    new DocsError('INVALID_XML', `${name} is not well-formed XML: ${message}`)
  const parser = new DOMParser({
    onError(level, message) {
      if (level !== 'warning') {
        throw invalid(message)
      }
    }
  })
  let doc: Document
  try {
    doc = parser.parseFromString(text, 'application/xml')
  } catch (err) {
    if (err instanceof DocsError) {
      throw err
    }

    const cause = causeOf(err)
    if (cause instanceof DocsError) {
      throw cause
    }

    throw invalid(errorMessage(err))
  }

  if (!doc.documentElement) {
    throw new DocsError('INVALID_XML', `${name} has no root element`)
  }

  return doc
}

export function toString(doc: Document): string {
  return new XMLSerializer().serializeToString(doc)
}

export function serialize(doc: Document): Uint8Array {
  const text = toString(doc)
  if (!text.startsWith('<?xml')) {
    return utf8.encode(DECLARATION + text)
  }

  return utf8.encode(text.replace(/^(<\?xml[^?]*?encoding=)(["'])[^"']*\2/i, '$1$2UTF-8$2'))
}

export function root(doc: Document): Element {
  const el = doc.documentElement
  if (!el) {
    throw new DocsError('INVALID_XML', 'the part has no root element')
  }

  return el
}

export function owner(node: Node): Document {
  const doc = node.ownerDocument
  if (!doc) {
    throw new DocsError('INVALID_XML', 'the element belongs to no document')
  }

  return doc
}

export function isElement(node: Node | null, namespace?: string, local?: string): node is Element {
  return (
    node instanceof Element &&
    (namespace === undefined || node.namespaceURI === namespace) &&
    (local === undefined || node.localName === local)
  )
}

export function children(el: Node, namespace?: string, local?: string): Element[] {
  const out: Element[] = []
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (isElement(n, namespace, local)) {
      out.push(n)
    }
  }

  return out
}

export function child(el: Node, namespace?: string, local?: string): Element | null {
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (isElement(n, namespace, local)) return n
  }

  return null
}

export function descendants(el: Element | Document, namespace: string, local: string): Element[] {
  return Array.from(el.getElementsByTagNameNS(namespace, local))
}

export function text(node: Node | null): string {
  return node?.textContent ?? ''
}

export function clean(value: string): string {
  return value.replace(NOT_XML, '')
}

export function setText(el: Element, value: string): void {
  while (el.firstChild) {
    el.removeChild(el.firstChild)
  }

  el.appendChild(owner(el).createTextNode(clean(value)))
}

export function attr(el: Element, name: string): string {
  return el.getAttribute(name) ?? ''
}

export function setAttributes(el: Element, attributes: Attributes = {}): Element {
  for (const [key, value] of Object.entries(attributes)) {
    if (value !== undefined && value !== null) {
      el.setAttribute(key, String(value))
    }
  }

  return el
}

export function named(doc: Document, qualifiedName: string, namespace: string): Element {
  return doc.createElementNS(namespace, qualifiedName)
}

export function qualify(doc: Document, namespace: string, local: string): string {
  const top = root(doc)
  if (top.namespaceURI === namespace && !top.prefix) {
    return local
  }

  const prefix = top.lookupPrefix(namespace) ?? PREFIXES.get(namespace)

  return prefix ? `${prefix}:${local}` : local
}

export function builder(doc: Document, namespace: string): Build {
  return (local, attributes) =>
    setAttributes(named(doc, qualify(doc, namespace, local), namespace), attributes)
}

export interface QualifiedName {
  namespace: string
  local: string
}

export function setNsAttribute(el: Element, name: QualifiedName, value: string): void {
  const doc = owner(el)
  const prefix =
    root(doc).lookupPrefix(name.namespace) ?? PREFIXES.get(name.namespace) ?? name.local
  el.setAttributeNS(name.namespace, `${prefix}:${name.local}`, value)
}

export function insertOrdered(parent: Element, el: Element, order: readonly string[]): Element {
  const rank = order.indexOf(el.localName ?? '')
  for (let n = parent.firstChild; n; n = n.nextSibling) {
    if (!(n instanceof Element)) continue
    const other = order.indexOf(n.localName ?? '')
    if (other > rank || other === -1) {
      parent.insertBefore(el, n)
      return el
    }
  }

  parent.appendChild(el)

  return el
}

export function remove(node: Node): void {
  node.parentNode?.removeChild(node)
}

export function append(parent: Node, node: Element): Element {
  parent.appendChild(node)
  return node
}

export function clone(el: Element): Element {
  const copy = el.cloneNode(true)
  if (!(copy instanceof Element)) {
    throw new DocsError('INVALID_XML', 'cloning an element produced no element')
  }

  return copy
}
