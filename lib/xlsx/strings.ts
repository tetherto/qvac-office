import * as xml from '../xml.ts'
import type { Element } from '../xml.ts'
import type { Package } from '../opc/package.ts'

const X = xml.ns.x
const TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
const XML_NS = 'http://www.w3.org/XML/1998/namespace'
const ESCAPE = /_x([0-9A-Fa-f]{4})_/g
const LITERAL = /_(?=x[0-9A-Fa-f]{4}_)/g
const UNWRITABLE =
  /[\u0000-\u0008\u000B-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g

function decodeText(text: string): string {
  return text.replace(ESCAPE, (_, hex: string) => String.fromCharCode(parseInt(hex, 16)))
}

function encodeText(text: string): string {
  return text
    .replace(LITERAL, '_x005F_')
    .replace(
      UNWRITABLE,
      (ch) => `_x${ch.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`
    )
}

export function runText(container: Element): string {
  const t = xml.child(container, X, 't')
  if (t) {
    return decodeText(xml.text(t))
  }

  return xml
    .children(container, X, 'r')
    .map((r) => decodeText(xml.text(xml.child(r, X, 't'))))
    .join('')
}

export class SharedStrings {
  package: Package
  workbookPart: string
  part: string | null
  private items: string[] | null
  private index: Map<string, number>

  constructor(pkg: Package, workbookPart: string) {
    this.package = pkg
    this.workbookPart = workbookPart
    this.part = pkg.related(workbookPart, REL)[0] ?? null
    this.items = null
    this.index = new Map()
  }

  get(i: number): string | null {
    return this.load()[i] ?? null
  }

  add(value: string): number {
    const items = this.load()
    const known = this.index.get(value)
    if (known !== undefined) {
      return known
    }

    const part = this.ensurePart()
    const doc = this.package.xml(part)
    const top = xml.root(doc)
    const si = xml.named(doc, 'si', X)
    const t = xml.append(si, xml.named(doc, 't', X))
    if (/^\s|\s$/.test(value)) {
      t.setAttributeNS(XML_NS, 'xml:space', 'preserve')
    }

    t.appendChild(doc.createTextNode(encodeText(value)))
    top.appendChild(si)
    const i = items.length
    items.push(value)
    this.index.set(value, i)
    top.setAttribute('uniqueCount', String(items.length))
    top.removeAttribute('count')
    this.package.touch(part)

    return i
  }

  private ensurePart(): string {
    if (this.part) return this.part
    const part = '/xl/sharedStrings.xml'
    this.package.add(part, xml.parse(`<sst xmlns="${X}"/>`), TYPE)
    this.package.rels(this.workbookPart).add(REL, part)
    this.part = part
    return part
  }

  private load(): string[] {
    if (this.items) return this.items
    const items: string[] = []
    this.items = items
    if (!this.part) {
      return items
    }

    for (const si of xml.children(this.package.root(this.part), X, 'si')) {
      const value = runText(si)
      if (!this.index.has(value)) {
        this.index.set(value, items.length)
      }

      items.push(value)
    }

    return items
  }
}
