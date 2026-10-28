import * as xml from '../xml.ts'
import { DocsError } from '../errors.ts'
import type { Relationships } from '../opc/relationships.ts'

const HYPERLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const ALLOWED = /^(?:https?:\/\/[^\s<>"{}|\\^`]+|mailto:[^\s<>"{}|\\^`]+@[^\s<>"{}|\\^`]+)$/i
const MAX_LENGTH = 2048

export function validateUrl(url: unknown): string {
  if (typeof url !== 'string' || url.length > MAX_LENGTH || !ALLOWED.test(url)) {
    throw new DocsError(
      'INVALID_LINK',
      `a link must be an http, https or mailto URL of at most ${MAX_LENGTH} characters, got ${JSON.stringify(url)}`,
      { field: 'url', expected: 'http://, https:// or mailto: URL', got: url }
    )
  }

  return url
}

export function addHyperlink(rels: Relationships, url: string): string {
  const taken = new Set(rels.all().map((rel) => rel.id))
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
