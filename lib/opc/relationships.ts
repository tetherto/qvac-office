import * as xml from '../xml.ts'
import * as paths from './paths.ts'
import type { Document, Element } from '../xml.ts'

export interface Relationship {
  id: string
  type: string
  target: string
  external: boolean
  part: string | null
}

export class Relationships {
  source: string
  name: string
  doc: Document
  dirty: boolean

  constructor(source: string, bytes: Uint8Array | null) {
    this.source = source
    this.name = paths.relsFor(source)
    this.doc = xml.parse(bytes ?? `<Relationships xmlns="${xml.ns.rels}"/>`, this.name)
    this.dirty = false
  }

  get root(): Element {
    return xml.root(this.doc)
  }

  all(): Relationship[] {
    return xml.children(this.root, xml.ns.rels, 'Relationship').map((el) => {
      const external = el.getAttribute('TargetMode') === 'External'
      const target = xml.attr(el, 'Target')
      return {
        id: xml.attr(el, 'Id'),
        type: xml.attr(el, 'Type'),
        target,
        external,
        part: external || !target ? null : paths.resolve(this.source, target)
      }
    })
  }

  get(id: string | null): Relationship | null {
    return this.all().find((rel) => rel.id === id) ?? null
  }

  byType(type: string): Relationship[] {
    return this.all().filter((rel) => rel.type === type)
  }

  targeting(partName: string): Relationship[] {
    return this.all().filter((rel) => rel.part === partName)
  }

  add(type: string, partName: string): string {
    const taken = new Set(this.all().map((rel) => rel.id))
    let n = taken.size + 1
    while (taken.has(`rId${n}`)) {
      n++
    }

    const id = `rId${n}`
    const el = xml.named(this.doc, 'Relationship', xml.ns.rels)
    this.root.appendChild(
      xml.setAttributes(el, { Id: id, Type: type, Target: paths.relative(this.source, partName) })
    )
    this.dirty = true

    return id
  }

  remove(id: string): boolean {
    const el = xml
      .children(this.root, xml.ns.rels, 'Relationship')
      .find((r) => r.getAttribute('Id') === id)
    if (!el) {
      return false
    }

    xml.remove(el)
    this.dirty = true

    return true
  }

  serialize(): Uint8Array {
    return xml.serialize(this.doc)
  }
}
