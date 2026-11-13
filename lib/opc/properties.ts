import * as xml from '../xml.ts'
import { DocsError } from '../errors.ts'
import type { Document, Element } from '../xml.ts'
import type { Package } from './package.ts'

export interface SaveOptions {
  modified?: Date
  modifiedBy?: string
}

const CORE_REL =
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
const APP_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'
const CORE_TYPE = 'application/vnd.openxmlformats-package.core-properties+xml'
const CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
const DC = 'http://purl.org/dc/elements/1.1/'
const DCTERMS = 'http://purl.org/dc/terms/'
const XSI = 'http://www.w3.org/2001/XMLSchema-instance'
const EP = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'

interface Field {
  namespace: string
  name: string
}

function setField(doc: Document, field: Field, value: string): Element {
  const top = xml.root(doc)
  const local = field.name.split(':')[1]
  const el =
    xml.child(top, field.namespace, local) ??
    xml.append(top, xml.named(doc, field.name, field.namespace))
  xml.setText(el, value)

  return el
}

function corePart(pkg: Package): string {
  const existing = pkg.related('/', CORE_REL)[0]
  if (existing) {
    return existing
  }

  const part = '/docProps/core.xml'
  pkg.add(part, xml.parse(`<cp:coreProperties xmlns:cp="${CP}" xmlns:dc="${DC}"/>`), CORE_TYPE)
  pkg.rels('/').add(CORE_REL, part)

  return part
}

export function applySaveOptions(pkg: Package, opts: SaveOptions = {}): void {
  const { modified, modifiedBy } = opts
  if (modified === undefined && modifiedBy === undefined) {
    return
  }

  if (modified !== undefined && !(modified instanceof Date && !Number.isNaN(modified.getTime()))) {
    throw new DocsError('INVALID_OPTIONS', 'modified must be a valid Date')
  }

  if (modifiedBy !== undefined && typeof modifiedBy !== 'string') {
    throw new DocsError('INVALID_OPTIONS', 'modifiedBy must be text')
  }

  const part = corePart(pkg)
  const doc = pkg.xml(part)
  if (modifiedBy !== undefined) {
    setField(doc, { namespace: CP, name: 'cp:lastModifiedBy' }, modifiedBy)
  }

  if (modified !== undefined) {
    const stamp = modified.toISOString().replace(/\.\d{3}Z$/, 'Z')
    const el = setField(doc, { namespace: DCTERMS, name: 'dcterms:modified' }, stamp)
    el.setAttributeNS(XSI, 'xsi:type', 'dcterms:W3CDTF')
  }

  pkg.touch(part)
}

export function refreshSummary(pkg: Package, counts: Record<string, number> = {}): void {
  const part = pkg.related('/', APP_REL)[0]
  if (!part) {
    return
  }

  const top = pkg.root(part)
  let changed = false
  for (const local of ['HeadingPairs', 'TitlesOfParts']) {
    for (const el of xml.children(top, EP, local)) {
      xml.remove(el)
      changed = true
    }
  }

  for (const [local, value] of Object.entries(counts)) {
    const el = xml.child(top, EP, local)
    if (el && xml.text(el) !== String(value)) {
      xml.setText(el, String(value))
      changed = true
    }
  }

  if (changed) {
    pkg.touch(part)
  }
}
