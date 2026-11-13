import * as xml from '../xml.ts'
import * as text from '../drawingml/text.ts'
import * as chart from '../drawingml/chart.ts'
import { DocsError } from '../errors.ts'
import type { Document, Element } from '../xml.ts'
import type { Box } from './inherit.ts'

export type ShapeKind =
  'text' | 'shape' | 'picture' | 'group' | 'connector' | 'table' | 'chart' | 'diagram' | 'frame'

export interface ShapeEntry {
  el: Element
  id: number | null
  name: string
  kind: ShapeKind
  placeholder: string | null
  group: number | null
}

export type Cell = string | number | null

export interface PictureFrame {
  id: number
  rId: string
  box: Box
  name?: string
  description?: string
}

export interface TextBoxFrame {
  id: number
  box: Box
  text?: unknown
  style?: unknown
  fill?: unknown
  name?: string
}

export interface TableFrame {
  id: number
  box: Box
  rows: unknown
  header?: boolean
  name?: string
  tableStyle: string
}

export interface ChartFrame {
  id: number
  box: Box
  rId: string
  name?: string
}

const { a: A, p: P, r: R } = xml.ns

export const SHAPE_ORDER = ['nvSpPr', 'spPr', 'style', 'txBody', 'extLst']
const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table'

function graphicKind(el: Element): ShapeKind {
  if (xml.descendants(el, A, 'tbl').length) return 'table'
  const data = xml.descendants(el, A, 'graphicData')[0]
  const uri = data?.getAttribute('uri') ?? ''
  if (uri.endsWith('/chart')) {
    return 'chart'
  }

  if (uri.endsWith('/diagram')) {
    return 'diagram'
  }

  return 'frame'
}

function kindOf(el: Element): ShapeKind | null {
  switch (el.localName) {
    case 'sp': {
      const nv = xml.children(el)[0]
      const placeholder = nv ? xml.descendants(nv, P, 'ph').length > 0 : false
      return xml.child(el, P, 'txBody') || placeholder ? 'text' : 'shape'
    }

    case 'pic':
      return 'picture'
    case 'grpSp':
      return 'group'
    case 'cxnSp':
      return 'connector'
    case 'graphicFrame':
      return graphicKind(el)
    default:
      return null
  }
}

function entryOf(el: Element, kind: ShapeKind, group: number | null): ShapeEntry {
  const nv = xml.children(el)[0]
  const cNvPr = nv ? xml.child(nv, P, 'cNvPr') : null
  const nvPr = nv ? xml.child(nv, P, 'nvPr') : null
  const ph = nvPr ? xml.child(nvPr, P, 'ph') : null
  return {
    el,
    id: cNvPr ? Number(cNvPr.getAttribute('id')) : null,
    name: cNvPr?.getAttribute('name') ?? '',
    kind,
    placeholder: ph ? ph.getAttribute('type') || 'obj' : null,
    group
  }
}

function walk(container: Element, group: number | null, out: ShapeEntry[]): void {
  for (const el of xml.children(container, P)) {
    const kind = kindOf(el)
    if (!kind) {
      continue
    }

    const shape = entryOf(el, kind, group)
    out.push(shape)
    if (kind === 'group') {
      walk(el, shape.id, out)
    }
  }
}

export function shapesOf(doc: Document): ShapeEntry[] {
  const spTree = xml.descendants(doc, P, 'spTree')[0]
  const out: ShapeEntry[] = []
  if (spTree) {
    walk(spTree, null, out)
  }

  return out
}

export function placeholderFrom(doc: Document, source: ShapeEntry, id: number): Element {
  const p = xml.builder(doc, P)
  const a = xml.builder(doc, A)
  const sourceNv = xml.children(source.el)[0]
  const sourceNvPr = sourceNv ? xml.child(sourceNv, P, 'nvPr') : null
  const sourcePh = sourceNvPr ? xml.child(sourceNvPr, P, 'ph') : null
  const sp = p('sp')
  const nvSpPr = xml.append(sp, p('nvSpPr'))
  nvSpPr.appendChild(p('cNvPr', { id, name: source.name || `Placeholder ${id - 1}` }))
  xml.append(nvSpPr, p('cNvSpPr')).appendChild(a('spLocks', { noGrp: 1 }))
  const ph = xml.append(xml.append(nvSpPr, p('nvPr')), p('ph'))
  for (const name of ['type', 'orient', 'sz', 'idx']) {
    const value = sourcePh?.getAttribute(name)
    if (value !== null && value !== undefined) {
      ph.setAttribute(name, value)
    }
  }

  sp.appendChild(p('spPr'))
  text.ensureTextBody(sp, SHAPE_ORDER)

  return sp
}

function xfrm(doc: Document, box: Box, namespace: string): Element {
  const a = xml.builder(doc, A)
  const el = xml.builder(doc, namespace)('xfrm')
  el.appendChild(a('off', { x: box.x, y: box.y }))
  el.appendChild(a('ext', { cx: box.width, cy: box.height }))
  return el
}

function rectangle(doc: Document): Element {
  const a = xml.builder(doc, A)
  const geom = a('prstGeom', { prst: 'rect' })
  geom.appendChild(a('avLst'))
  return geom
}

export function pictureElement(doc: Document, frame: PictureFrame): Element {
  const p = xml.builder(doc, P)
  const a = xml.builder(doc, A)
  const pic = p('pic')
  const nv = xml.append(pic, p('nvPicPr'))
  nv.appendChild(
    p('cNvPr', {
      id: frame.id,
      name: frame.name || `Picture ${frame.id - 1}`,
      descr: frame.description
    })
  )
  xml.append(nv, p('cNvPicPr')).appendChild(a('picLocks', { noChangeAspect: 1 }))
  nv.appendChild(p('nvPr'))
  const fill = xml.append(pic, p('blipFill'))
  const blip = xml.append(fill, a('blip'))
  xml.setNsAttribute(blip, { namespace: R, local: 'embed' }, frame.rId)
  xml.append(fill, a('stretch')).appendChild(a('fillRect'))
  const spPr = xml.append(pic, p('spPr'))
  spPr.appendChild(xfrm(doc, frame.box, A))
  spPr.appendChild(rectangle(doc))

  return pic
}

export function nextShapeId(doc: Document): number {
  let max = 0
  for (const el of xml.descendants(doc, P, 'cNvPr')) {
    max = Math.max(max, Number(el.getAttribute('id')) || 0)
  }

  return max + 1
}

function solidFill(doc: Document, color: string): Element {
  const a = xml.builder(doc, A)
  const fill = a('solidFill')
  fill.appendChild(a('srgbClr', { val: color.replace('#', '').toUpperCase() }))
  return fill
}

export function validateColor(color: unknown, name = 'color'): string {
  if (typeof color !== 'string' || !/^#?[0-9a-fA-F]{6}$/.test(color)) {
    throw new DocsError(
      'INVALID_STYLE',
      `${name} must be a hex RGB value like "1A73E8", got ${String(color)}`
    )
  }

  return color
}

export function textBoxElement(doc: Document, frame: TextBoxFrame): Element {
  const p = xml.builder(doc, P)
  const a = xml.builder(doc, A)
  const sp = p('sp')
  const nv = xml.append(sp, p('nvSpPr'))
  nv.appendChild(p('cNvPr', { id: frame.id, name: frame.name || `TextBox ${frame.id - 1}` }))
  nv.appendChild(p('cNvSpPr', { txBox: 1 }))
  nv.appendChild(p('nvPr'))
  const spPr = xml.append(sp, p('spPr'))
  spPr.appendChild(xfrm(doc, frame.box, A))
  spPr.appendChild(rectangle(doc))
  spPr.appendChild(frame.fill ? solidFill(doc, validateColor(frame.fill, 'fill')) : a('noFill'))
  const txBody = text.ensureTextBody(sp, SHAPE_ORDER)
  xml.child(txBody, A, 'bodyPr')?.setAttribute('wrap', 'square')
  text.writeParagraphs(txBody, frame.text === undefined ? '' : frame.text)
  if (frame.style) {
    text.applyStyle(txBody, frame.style)
  }

  return sp
}

function isCell(value: unknown): value is Cell {
  return (
    value === null ||
    typeof value === 'string' ||
    (typeof value === 'number' && Number.isFinite(value))
  )
}

export function validateRows(rows: unknown): Cell[][] {
  if (!Array.isArray(rows) || rows.length === 0 || !rows.every(Array.isArray)) {
    throw new DocsError(
      'INVALID_TABLE',
      'rows must be a non-empty array of rows, each an array of cells'
    )
  }

  const out: Cell[][] = rows.map((row: unknown[], r) =>
    row.map((value, c) => {
      if (isCell(value)) return value
      throw new DocsError('INVALID_TABLE', `cell ${r},${c} must be a string, a number or null`)
    })
  )
  const columns = Math.max(...out.map((r) => r.length))
  if (columns === 0) {
    throw new DocsError('INVALID_TABLE', 'the table has no columns')
  }

  if (out.length > 1000 || columns > 64) {
    throw new DocsError('INVALID_TABLE', 'a table holds at most 1000 rows and 64 columns')
  }

  return out
}

interface GraphicSpec {
  id: number
  box: Box
  name: string
  uri: string
  locked: boolean
}

interface GraphicParts {
  frame: Element
  data: Element
}

function graphicFrame(doc: Document, spec: GraphicSpec): GraphicParts {
  const p = xml.builder(doc, P)
  const a = xml.builder(doc, A)
  const frame = p('graphicFrame')
  const nv = xml.append(frame, p('nvGraphicFramePr'))
  nv.appendChild(p('cNvPr', { id: spec.id, name: spec.name }))
  const locks = xml.append(nv, p('cNvGraphicFramePr'))
  if (spec.locked) {
    locks.appendChild(a('graphicFrameLocks', { noGrp: 1 }))
  }

  nv.appendChild(p('nvPr'))
  frame.appendChild(xfrm(doc, spec.box, P))
  const data = xml.append(xml.append(frame, a('graphic')), a('graphicData', { uri: spec.uri }))

  return { frame, data }
}

function tableCell(doc: Document, value: Cell | undefined): Element {
  const a = xml.builder(doc, A)
  const tc = a('tc')
  const txBody = xml.append(tc, a('txBody'))
  for (const local of ['bodyPr', 'lstStyle', 'p']) {
    txBody.appendChild(a(local))
  }

  text.writeParagraphs(txBody, value === null || value === undefined ? '' : String(value))
  tc.appendChild(a('tcPr'))

  return tc
}

export function tableElement(doc: Document, frame: TableFrame): Element {
  const rows = validateRows(frame.rows)
  const columns = Math.max(...rows.map((r) => r.length))
  const a = xml.builder(doc, A)
  const { frame: el, data } = graphicFrame(doc, {
    id: frame.id,
    box: frame.box,
    name: frame.name || `Table ${frame.id - 1}`,
    uri: TABLE_URI,
    locked: true
  })
  const tbl = xml.append(data, a('tbl'))
  const tblPr = xml.append(
    tbl,
    a('tblPr', { firstRow: frame.header === false ? undefined : 1, bandRow: 1 })
  )
  xml.setText(xml.append(tblPr, a('tableStyleId')), frame.tableStyle)
  const grid = xml.append(tbl, a('tblGrid'))
  const width = Math.floor(frame.box.width / columns)
  for (let c = 0; c < columns; c++) {
    grid.appendChild(a('gridCol', { w: width }))
  }

  const height = Math.floor(frame.box.height / rows.length)
  for (const row of rows) {
    const tr = xml.append(tbl, a('tr', { h: height }))
    for (let c = 0; c < columns; c++) {
      tr.appendChild(tableCell(doc, row[c]))
    }
  }

  return el
}

export function chartFrameElement(doc: Document, frame: ChartFrame): Element {
  const { frame: el, data } = graphicFrame(doc, {
    id: frame.id,
    box: frame.box,
    name: frame.name || `Chart ${frame.id - 1}`,
    uri: chart.GRAPHIC_URI,
    locked: false
  })
  const ref = xml.append(data, xml.builder(doc, chart.C)('chart'))
  xml.setNsAttribute(ref, { namespace: R, local: 'id' }, frame.rId)

  return el
}

export function setBackground(doc: Document, color: unknown): void {
  const cSld = xml.child(xml.root(doc), P, 'cSld')
  if (!cSld) {
    throw new DocsError('BROKEN_PPTX', 'the slide has no common slide data')
  }

  const existing = xml.child(cSld, P, 'bg')
  if (existing) {
    xml.remove(existing)
  }

  if (color === null) {
    return
  }

  const p = xml.builder(doc, P)
  const bg = p('bg')
  const bgPr = xml.append(bg, p('bgPr'))
  bgPr.appendChild(solidFill(doc, validateColor(color)))
  bgPr.appendChild(xml.builder(doc, A)('effectLst'))
  cSld.insertBefore(bg, cSld.firstChild)
}
