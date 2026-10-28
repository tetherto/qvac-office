import * as xml from '../xml.ts'
import { extension } from './paths.ts'
import type { Document, Element } from '../xml.ts'

export const CONTENT_TYPES_NAME = '/[Content_Types].xml'

export class ContentTypes {
  doc: Document
  dirty: boolean
  private original: Uint8Array | null

  constructor(bytes: Uint8Array | null) {
    this.original = bytes
    this.doc = xml.parse(bytes ?? `<Types xmlns="${xml.ns.ct}"/>`, CONTENT_TYPES_NAME)
    this.dirty = !bytes
  }

  get root(): Element {
    return xml.root(this.doc)
  }

  get(partName: string): string | null {
    const override = this.override(partName)
    if (override) {
      return override.getAttribute('ContentType')
    }

    const ext = extension(partName)
    const entry = this.defaults().find((el) => xml.attr(el, 'Extension').toLowerCase() === ext)

    return entry ? entry.getAttribute('ContentType') : null
  }

  hasDefault(ext: string): boolean {
    return this.defaults().some(
      (el) => xml.attr(el, 'Extension').toLowerCase() === ext.toLowerCase()
    )
  }

  addDefault(ext: string, contentType: string): void {
    if (this.hasDefault(ext)) return
    const el = xml.setAttributes(xml.named(this.doc, 'Default', xml.ns.ct), {
      Extension: ext,
      ContentType: contentType
    })
    this.root.insertBefore(el, xml.child(this.root, xml.ns.ct, 'Override'))
    this.dirty = true
  }

  setOverride(partName: string, contentType: string): void {
    const existing = this.override(partName)
    if (existing?.getAttribute('ContentType') === contentType) {
      return
    }

    if (existing) {
      existing.setAttribute('ContentType', contentType)
    } else {
      const el = xml.named(this.doc, 'Override', xml.ns.ct)
      this.root.appendChild(xml.setAttributes(el, { PartName: partName, ContentType: contentType }))
    }

    this.dirty = true
  }

  removeOverride(partName: string): void {
    const existing = this.override(partName)
    if (!existing) {
      return
    }

    xml.remove(existing)
    this.dirty = true
  }

  serialize(): Uint8Array {
    if (!this.dirty && this.original) return this.original
    return xml.serialize(this.doc)
  }

  private defaults(): Element[] {
    return xml.children(this.root, xml.ns.ct, 'Default')
  }

  private override(partName: string): Element | undefined {
    const wanted = partName.toLowerCase()
    return xml
      .children(this.root, xml.ns.ct, 'Override')
      .find((el) => xml.attr(el, 'PartName').toLowerCase() === wanted)
  }
}
