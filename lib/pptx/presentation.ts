import { Package, type OpenOptions } from '../opc/package.ts'
import * as xml from '../xml.ts'
import * as text from '../drawingml/text.ts'
import * as chart from '../drawingml/chart.ts'
import * as media from '../opc/media.ts'
import * as properties from '../opc/properties.ts'
import { probe, type ImageInfo } from '../image.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import * as inherit from './inherit.ts'
import * as fit from './fit.ts'
import * as template from './template.ts'
import * as slides from './slides.ts'
import * as notes from './notes.ts'
import {
  fitted,
  framed,
  pictureArea,
  type Area,
  type Beside,
  type Placed,
  type Role
} from './picture-area.ts'
import { pick, describe, layoutName, layoutType, BODY_TYPES } from './read.ts'
import {
  SHAPE_ORDER,
  shapesOf,
  placeholderFrom,
  pictureElement,
  textBoxElement,
  tableElement,
  chartFrameElement,
  setBackground,
  nextShapeId,
  type ShapeEntry
} from './shapes.ts'
import type { Document, Element } from '../xml.ts'
import type { Box } from './inherit.ts'
import type { SaveOptions } from '../opc/properties.ts'
import type { CreateOptions, SlideSize } from './template.ts'
import type { ReadParagraph, Paragraph, TextInput, Style } from '../drawingml/text.ts'
import type { DataResult, ChartData, Changed as ChartChanged } from '../drawingml/chart.ts'
import type {
  ShapeRef,
  CellRef,
  ReadOptions,
  SlideInfo,
  LayoutInfo,
  TextResult,
  FitOptions,
  StyleOptions,
  RunCount,
  AddSlideOptions,
  AddedSlide,
  Placement,
  PictureOptions,
  TextBoxOptions,
  TableOptions,
  ChartOptions,
  PlacedShape,
  ReplaceOptions,
  ReplacedPicture,
  MoveOptions,
  Moved,
  Deleted,
  Remaining,
  Changed,
  ShapeInfo
} from './types.ts'
import * as units from '../units.ts'

interface Located {
  entry: slides.SlideEntry
  found: ShapeEntry
}

interface NewSlide {
  doc: Document
  layout: string
  at: number
  ids: string[]
  before: Element | null
}

interface Written {
  part: string
  txBody: Element
}

const { a: A, p: P, r: R, p14: P14 } = xml.ns
const STRICT_REL = 'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument'
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
const REL = {
  officeDocument: REL_BASE + 'officeDocument',
  slide: REL_BASE + 'slide',
  notesSlide: REL_BASE + 'notesSlide',
  slideLayout: REL_BASE + 'slideLayout',
  image: REL_BASE + 'image',
  chart: REL_BASE + 'chart',
  hyperlink: REL_BASE + 'hyperlink'
}
const SLIDE_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
const MAIN_TYPES = new Set([
  'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml',
  'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml',
  'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml',
  'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml',
  'application/vnd.ms-powerpoint.template.macroEnabled.main+xml'
])
const SKIPPED_PLACEHOLDERS = new Set(['dt', 'ftr', 'sldNum'])
const TITLE_TYPES = new Set(['title', 'ctrTitle'])
const SVG_BLIP_EXT = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}'
const BOX_KEYS = ['x', 'y', 'width', 'height'] as const
const EMU_PER_PIXEL = 9525
const EMPTY_SLIDE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`

function checkPlacement(opts: Placement): void {
  for (const key of BOX_KEYS) {
    const value: unknown = opts[key]
    if (
      value !== undefined &&
      !(typeof value === 'number' && Number.isInteger(value) && value >= 0)
    ) {
      throw new DocsError(
        'INVALID_BOX',
        `${key} must be a non-negative whole number of EMU, got ${String(value)}`
      )
    }
  }
}

function withLevels(paragraphs: Paragraph[], before: ReadParagraph[]): ReadParagraph[] {
  const list = paragraphs.length === 0 ? [{ text: '' }] : paragraphs
  return list.map((p, i) => ({
    text: p.text,
    level: p.level ?? (before[Math.min(i, before.length - 1)] ?? { level: 0 }).level
  }))
}

function sameParagraphs(a: ReadParagraph[], b: ReadParagraph[]): boolean {
  return a.length === b.length && a.every((p, i) => p.text === b[i].text && p.level === b[i].level)
}

function applyAutofit(txBody: Element, estimate: fit.Estimate): void {
  const a = xml.builder(xml.owner(txBody), A)
  let bodyPr = xml.child(txBody, A, 'bodyPr')
  if (!bodyPr) {
    bodyPr = a('bodyPr')
    txBody.insertBefore(bodyPr, txBody.firstChild)
  }

  for (const local of ['noAutofit', 'spAutoFit', 'normAutofit']) {
    for (const el of xml.children(bodyPr, A, local)) {
      xml.remove(el)
    }
  }

  const autofit = a('normAutofit', {
    fontScale: estimate.fontScale < 1 ? Math.round(estimate.fontScale * 100000) : undefined,
    lnSpcReduction:
      estimate.lineReduction > 0 ? Math.round(estimate.lineReduction * 100000) : undefined
  })
  const after = xml
    .children(bodyPr, A)
    .find((el) => ['scene3d', 'sp3d', 'flatTx', 'extLst'].includes(el.localName ?? ''))
  bodyPr.insertBefore(autofit, after ?? null)
}

function textBodies(found: ShapeEntry): Element[] {
  if (found.kind === 'table') {
    return xml.descendants(found.el, A, 'txBody')
  }

  const own = xml.child(found.el, P, 'txBody')

  return own ? [own] : []
}

function findShape(shapes: ShapeEntry[], ref: unknown): ShapeEntry | undefined {
  if (ref === 'title') {
    return shapes.find((s) => s.placeholder !== null && TITLE_TYPES.has(s.placeholder))
  }

  if (ref === 'subtitle') {
    return shapes.find((s) => s.placeholder === 'subTitle')
  }

  if (ref === 'body') {
    return shapes.find(
      (s) => s.placeholder !== null && BODY_TYPES.has(s.placeholder) && s.el.localName === 'sp'
    )
  }

  return shapes.find((s) => String(s.id) === String(ref))
}

function dropSvgBlip(blip: Element): void {
  for (const ext of xml.descendants(blip, A, 'ext')) {
    if (ext.getAttribute('uri') === SVG_BLIP_EXT) {
      xml.remove(ext)
    }
  }

  const extLst = xml.child(blip, A, 'extLst')
  if (extLst && !xml.children(extLst).length) {
    xml.remove(extLst)
  }
}

function dropCrop(blipFill: Element): void {
  for (const srcRect of xml.children(blipFill, A, 'srcRect')) {
    xml.remove(srcRect)
  }

  for (const fillRect of xml.descendants(blipFill, A, 'fillRect')) {
    for (const side of ['l', 't', 'r', 'b']) {
      fillRect.removeAttribute(side)
    }
  }
}

function shapeNotFound(slide: unknown, ref: unknown, shapes: ShapeEntry[]): DocsError {
  const listing = shapes.map(
    (s) => `${s.id} ${s.kind}${s.placeholder ? ` (${s.placeholder})` : ''} "${s.name}"`
  )

  return new DocsError(
    'SHAPE_NOT_FOUND',
    `slide ${String(slide)} has no shape ${JSON.stringify(ref)}; its shapes are: ${listing.join(', ') || 'none'}`
  )
}

function slideFromLayout(layoutDoc: Document): Document {
  const doc = xml.parse(EMPTY_SLIDE)
  const spTree = xml.descendants(doc, P, 'spTree')[0]
  let nextId = 2
  for (const source of shapesOf(layoutDoc)) {
    const skip =
      source.el.localName !== 'sp' ||
      !source.placeholder ||
      SKIPPED_PLACEHOLDERS.has(source.placeholder)
    if (!skip) {
      spTree.appendChild(placeholderFrom(doc, source, nextId++))
    }
  }

  return doc
}

function checkSlideText(doc: Document, at: number, opts: AddSlideOptions): void {
  const shapes = shapesOf(doc)
  for (const [ref, value] of [
    ['title', opts.title],
    ['body', opts.body]
  ] as const) {
    if (value === undefined) continue
    text.normalizeParagraphs(value)
    if (!findShape(shapes, ref)) {
      throw shapeNotFound(at, ref, shapes)
    }
  }
}

interface SlideShape extends Placed {
  readonly el: Element
}

function roleOf(shape: ShapeEntry): Role {
  const { el, kind, placeholder } = shape
  if (kind === 'picture') {
    return 'picture'
  }

  if (el.localName !== 'sp' || !placeholder || !BODY_TYPES.has(placeholder)) {
    return 'other'
  }

  const txBody = xml.child(el, P, 'txBody')
  if (txBody && text.readParagraphs(txBody).some((p) => p.text.trim())) {
    return placeholder === 'subTitle' ? 'other' : 'body'
  }

  return placeholder === 'obj' ? 'slot' : 'blank'
}

function tableBody(tc: Element): Element {
  const existing = xml.child(tc, A, 'txBody')
  if (existing) {
    return existing
  }

  const a = xml.builder(xml.owner(tc), A)
  const txBody = a('txBody')
  for (const local of ['bodyPr', 'lstStyle', 'p']) {
    txBody.appendChild(a(local))
  }

  tc.insertBefore(txBody, tc.firstChild)

  return txBody
}

function ensureXfrm(el: Element): Element {
  const existing = inherit.xfrmOf(el)
  if (existing) {
    return existing
  }

  const doc = xml.owner(el)
  if (el.localName === 'graphicFrame') {
    const frameXfrm = xml.builder(doc, P)('xfrm')
    el.insertBefore(frameXfrm, xml.children(el)[1] ?? null)
    return frameXfrm
  }

  const propsName = el.localName === 'grpSp' ? 'grpSpPr' : 'spPr'
  let props = xml.child(el, P, propsName)
  if (!props) {
    props = xml.builder(doc, P)(propsName)
    el.insertBefore(props, xml.children(el)[1] ?? null)
  }

  const made = xml.builder(doc, A)('xfrm')
  props.insertBefore(made, props.firstChild)

  return made
}

function writeXfrm(el: Element, box: Box): void {
  const xfrm = ensureXfrm(el)
  const a = xml.builder(xml.owner(el), A)
  let off = xml.child(xfrm, A, 'off')
  if (!off) {
    off = a('off')
    xfrm.insertBefore(off, xfrm.firstChild)
  }

  let ext = xml.child(xfrm, A, 'ext')
  if (!ext) {
    ext = a('ext')
    xfrm.insertBefore(ext, off.nextSibling)
  }

  xml.setAttributes(off, { x: box.x, y: box.y })
  xml.setAttributes(ext, { cx: box.width, cy: box.height })
}

function checkBox(box: unknown): Partial<Box> {
  if (!isRecord(box)) {
    throw new DocsError('INVALID_BOX', 'box must be { x, y, width, height }')
  }

  const keys = Object.keys(box)
  if (!keys.length || keys.some((k) => !['x', 'y', 'width', 'height'].includes(k))) {
    throw new DocsError('INVALID_BOX', 'box takes x, y, width and height in EMU')
  }

  const out: Partial<Box> = {}
  for (const key of BOX_KEYS) {
    const value = box[key]
    if (value === undefined) {
      continue
    }

    if (!(typeof value === 'number' && Number.isInteger(value) && value >= 0)) {
      throw new DocsError(
        'INVALID_BOX',
        `${key} must be a non-negative whole number of EMU, got ${String(value)}`
      )
    }

    out[key] = value
  }

  return out
}

function shapeIds(el: Element): Set<string> {
  const own = el.localName === 'cNvPr' ? el : xml.descendants(el, P, 'cNvPr')[0]
  const all = [own, ...xml.descendants(el, P, 'cNvPr')].filter(
    (node): node is Element => node !== undefined
  )

  return new Set(all.map((node) => xml.attr(node, 'id')))
}

export class Presentation {
  package: Package
  main: string
  private slidesChanged: boolean
  private touchedSlides: Set<string>

  constructor(pkg: Package) {
    this.package = pkg
    this.slidesChanged = false
    this.touchedSlides = new Set()
    const main = pkg.related('/', REL.officeDocument)[0]
    if (!main || !MAIN_TYPES.has(pkg.contentType(main) ?? '')) {
      if (pkg.rels('/').byType(STRICT_REL).length) {
        throw new DocsError(
          'UNSUPPORTED_PPTX',
          'the presentation uses Strict Open XML, which is not supported'
        )
      }

      throw new DocsError('NOT_PPTX', 'the file is not a PowerPoint presentation')
    }

    this.main = main
    const top = pkg.root(main)
    if (top.namespaceURI !== P || top.localName !== 'presentation') {
      throw new DocsError(
        'UNSUPPORTED_PPTX',
        'the presentation uses Strict Open XML, which is not supported'
      )
    }
  }

  static open(bytes: Uint8Array, opts?: OpenOptions): Presentation {
    return new Presentation(Package.open(bytes, opts))
  }

  static create(opts: CreateOptions = {}): Presentation {
    const pkg = Package.create()
    template.build(pkg, opts)
    return new Presentation(pkg)
  }

  get changed(): boolean {
    return this.package.changed
  }

  get slideCount(): number {
    return slides.entries(this.deck).length
  }

  get slideSize(): SlideSize {
    const sz = xml.child(this.package.root(this.main), P, 'sldSz')
    return sz
      ? { width: Number(sz.getAttribute('cx')), height: Number(sz.getAttribute('cy')) }
      : { width: 9144000, height: 6858000 }
  }

  layouts(): LayoutInfo[] {
    return slides.allLayouts(this.deck).map(({ part }) => ({
      name: layoutName(this.package.xml(part)),
      type: layoutType(this.package.xml(part))
    }))
  }

  read(opts: ReadOptions = {}): SlideInfo[] {
    const all = slides.entries(this.deck)
    const wanted = pick(all.length, opts)
    return all.flatMap((entry, index) => (wanted(index) ? [this.slideInfo(entry, index)] : []))
  }

  setText(target: ShapeRef, value: TextInput, opts: FitOptions = {}): TextResult {
    if (opts.fit !== undefined && opts.fit !== 'shrink') {
      throw new DocsError('INVALID_OPTIONS', 'fit must be "shrink" or left out')
    }

    const { entry, found } = this.locate(target)
    if (found.kind === 'table') {
      throw new DocsError(
        'NOT_TEXT',
        `shape ${found.id} is a table; set its cells with setCellText`
      )
    }

    if (found.el.localName !== 'sp') {
      throw new DocsError('NOT_TEXT', `shape ${found.id} is a ${found.kind} and holds no text`)
    }

    const txBody = text.ensureTextBody(found.el, SHAPE_ORDER)
    const result = this.write({ part: entry.part, txBody }, value)
    const estimate = this.estimate(entry.part, found.el, txBody)
    if (!estimate) {
      return result
    }

    result.overflow = estimate.overflow
    if (opts.fit === 'shrink') {
      applyAutofit(txBody, estimate)
      this.package.touch(entry.part)
      result.fontScale = estimate.fontScale
    }

    return result
  }

  setCellText(cell: CellRef, value: TextInput): TextResult {
    const { entry, found } = this.locate(cell)
    if (found.kind !== 'table') {
      throw new DocsError('NOT_TABLE', `shape ${found.id} is a ${found.kind}, not a table`)
    }

    const { row, column } = cell
    const rows = xml.descendants(found.el, A, 'tr')
    const tr = rows[row]
    if (!tr) {
      throw new DocsError(
        'CELL_NOT_FOUND',
        `row ${row} does not exist; the table has ${rows.length} (0 to ${rows.length - 1})`
      )
    }

    const cells = xml.children(tr, A, 'tc')
    const tc = cells[column]
    if (!tc) {
      throw new DocsError(
        'CELL_NOT_FOUND',
        `column ${column} does not exist; the row has ${cells.length} (0 to ${cells.length - 1})`
      )
    }

    if (tc.getAttribute('hMerge') === '1' || tc.getAttribute('vMerge') === '1') {
      throw new DocsError(
        'CELL_MERGED',
        `cell ${row},${column} is covered by a merged cell; write the merge's top-left cell`
      )
    }

    return this.write({ part: entry.part, txBody: tableBody(tc) }, value)
  }

  setStyle(target: ShapeRef, style: Style, opts: StyleOptions = {}): RunCount {
    const { entry, found } = this.locate(target)
    const bodies = textBodies(found)
    if (bodies.length === 0) {
      throw new DocsError('NOT_TEXT', `shape ${found.id} is a ${found.kind} and holds no text`)
    }

    if (opts.paragraph !== undefined && bodies.length > 1) {
      throw new DocsError(
        'INVALID_STYLE',
        'paragraph cannot be used on a table; style the whole table instead'
      )
    }

    const runs = bodies.reduce((n, txBody) => n + text.applyStyle(txBody, style, opts), 0)
    this.package.touch(entry.part)

    return { runs }
  }

  format(target: ShapeRef, style: Style, opts?: StyleOptions): RunCount {
    return this.setStyle(target, style, opts)
  }

  addSlide(opts: AddSlideOptions = {}): AddedSlide {
    const all = slides.entries(this.deck)
    const at = opts.at ?? all.length
    if (!Number.isInteger(at) || at < 0 || at > all.length) {
      throw new DocsError(
        'SLIDE_NOT_FOUND',
        `at must be a position from 0 to ${all.length}, got ${String(opts.at)}`
      )
    }

    const layout = slides.pickLayout(this.deck, opts)
    const layoutDoc = this.package.xml(layout)
    const doc = slideFromLayout(layoutDoc)
    checkSlideText(doc, at, opts)
    if (opts.image !== undefined) {
      probe(opts.image)
    }

    this.insertSlide({ doc, layout, at, ids: all.map((e) => e.id), before: all[at]?.el ?? null })
    if (opts.title !== undefined) {
      this.setText({ slide: at, shape: 'title' }, opts.title)
    }

    if (opts.body !== undefined) {
      this.setText({ slide: at, shape: 'body' }, opts.body)
    }

    if (opts.image !== undefined) {
      this.addPicture(at, opts.image)
    }

    const shapes = shapesOf(this.package.xml(slides.slideAt(this.deck, at).part))

    return { index: at, layout: layoutName(layoutDoc), shapes: shapes.map(describe) }
  }

  addPicture(
    slide: number,
    bytes: Uint8Array,
    input: units.WithLengths<PictureOptions, units.BoxKey> = {}
  ): PlacedShape {
    const opts = units.lengths(input, units.BOX_KEYS) as PictureOptions
    const entry = slides.slideAt(this.deck, slide)
    const area = this.pictureArea(entry.part, opts)
    const box = this.pictureBox(probe(bytes), opts, area.box)
    const { rId } = media.addImage(this.package, bytes, { source: entry.part, dir: '../media' })
    const doc = this.package.xml(entry.part)
    const id = nextShapeId(doc)
    for (const shape of area.cleared) {
      xml.remove(shape.el)
    }

    if (area.beside) {
      this.narrow(entry.part, area.beside)
    }

    this.spTree(entry.part).appendChild(
      pictureElement(doc, { id, rId, box, name: opts.name, description: opts.description })
    )
    this.package.touch(entry.part)

    return { id, ...box }
  }

  addTextBox(
    slide: number,
    input: units.WithLengths<TextBoxOptions, units.BoxKey> = {}
  ): PlacedShape {
    const opts = units.lengths(input, units.BOX_KEYS) as TextBoxOptions
    const entry = slides.slideAt(this.deck, slide)
    const doc = this.package.xml(entry.part)
    const id = nextShapeId(doc)
    const box = this.box(opts)
    const el = textBoxElement(doc, {
      id,
      box,
      text: opts.text,
      style: opts.style,
      fill: opts.fill,
      name: opts.name
    })
    this.spTree(entry.part).appendChild(el)
    this.package.touch(entry.part)

    return { id, ...box }
  }

  addTable(slide: number, input: units.WithLengths<TableOptions, units.BoxKey>): PlacedShape {
    const opts = units.lengths(input, units.BOX_KEYS) as TableOptions
    const entry = slides.slideAt(this.deck, slide)
    const doc = this.package.xml(entry.part)
    const id = nextShapeId(doc)
    const box = this.box(opts)
    const el = tableElement(doc, {
      id,
      box,
      rows: opts.rows,
      header: opts.header,
      name: opts.name,
      tableStyle: template.TABLE_STYLE
    })
    this.spTree(entry.part).appendChild(el)
    this.package.touch(entry.part)

    return { id, ...box }
  }

  addChart(slide: number, input: units.WithLengths<ChartOptions, units.BoxKey>): PlacedShape {
    const spec = units.lengths(input, units.BOX_KEYS) as ChartOptions
    const entry = slides.slideAt(this.deck, slide)
    const box = this.box(spec)
    const part = this.package.freeName('/ppt/charts/chart%d.xml')
    chart.createChart({ pkg: this.package, part }, spec)
    const rId = this.package.rels(entry.part).add(REL.chart, part)
    const doc = this.package.xml(entry.part)
    const id = nextShapeId(doc)
    this.spTree(entry.part).appendChild(chartFrameElement(doc, { id, box, rId, name: spec.name }))
    this.package.touch(entry.part)
    return { id, ...box }
  }

  setChartData(target: ShapeRef, data: ChartData): DataResult {
    return chart.setChartData({ pkg: this.package, part: this.chartOf(target) }, data)
  }

  setChartTitle(target: ShapeRef, title: TextInput | null): ChartChanged {
    const value: unknown = title
    if (!(value === null || typeof value === 'string' || Array.isArray(value))) {
      throw new DocsError('INVALID_TEXT', 'title must be text, or null to remove it')
    }

    const ref = { pkg: this.package, part: this.chartOf(target) }
    if (value === null || typeof value === 'string') {
      return chart.setChartTitle(ref, value)
    }

    return chart.setChartTitle(ref, text.normalizeParagraphs(value))
  }

  setBackground(slide: number, color: string | null): Changed {
    const entry = slides.slideAt(this.deck, slide)
    setBackground(this.package.xml(entry.part), color)
    this.package.touch(entry.part)
    return { changed: true }
  }

  deleteShape(target: ShapeRef): Deleted {
    const { entry, found } = this.locate(target)
    const doc = this.package.xml(entry.part)
    const ids = shapeIds(found.el)
    const result: Deleted = { deleted: found.id }
    const timing = xml.child(xml.root(doc), P, 'timing')
    if (timing && xml.descendants(timing, P, 'spTgt').some((t) => ids.has(xml.attr(t, 'spid')))) {
      xml.remove(timing)
      result.removed = { animations: true }
    }

    for (const local of ['stCxn', 'endCxn']) {
      for (const link of xml.descendants(doc, A, local)) {
        if (ids.has(xml.attr(link, 'id'))) {
          xml.remove(link)
        }
      }
    }

    xml.remove(found.el)
    this.touchSlide(entry.part)

    return result
  }

  setBox(target: ShapeRef, raw: units.WithLengths<Partial<Box>, units.BoxKey>): PlacedShape {
    const input = units.lengths(raw, units.BOX_KEYS) as Partial<Box>
    const { entry, found } = this.locate(target)
    const box = checkBox(input)
    const current = inherit.box({ pkg: this.package, slide: entry.part, el: found.el })
    const merged = { ...current, ...box }
    const { x, y, width, height } = merged
    if (x === undefined || y === undefined || width === undefined || height === undefined) {
      throw new DocsError(
        'INVALID_BOX',
        `shape ${found.id} has no position to start from; give x, y, width and height`
      )
    }

    if (width === 0 || height === 0) {
      throw new DocsError('INVALID_BOX', 'the shape would have no visible size')
    }

    const next = { x, y, width, height }
    writeXfrm(found.el, next)
    this.touchSlide(entry.part)

    return { id: found.id ?? 0, ...next }
  }

  replacePicture(target: ShapeRef, bytes: Uint8Array, opts: ReplaceOptions = {}): ReplacedPicture {
    const { entry, found } = this.locate(target)
    if (found.kind !== 'picture') {
      throw new DocsError('NOT_PICTURE', `shape ${found.id} is a ${found.kind}, not a picture`)
    }

    const info = probe(bytes)
    const blip = xml.descendants(found.el, A, 'blip')[0]
    if (!blip) {
      throw new DocsError('NOT_PICTURE', `shape ${found.id} has no image to replace`)
    }

    const { rId } = media.addImage(this.package, bytes, { source: entry.part, dir: '../media' })
    xml.setNsAttribute(blip, { namespace: R, local: 'embed' }, rId)
    dropSvgBlip(blip)
    const blipFill = blip.parentNode
    if (xml.isElement(blipFill)) {
      dropCrop(blipFill)
    }

    const current = inherit.box({ pkg: this.package, slide: entry.part, el: found.el })
    let result: ReplacedPicture = { id: found.id }
    if (current) {
      const next = opts.stretch ? current : fitted(info, current)
      result = this.setBox({ slide: target.slide, shape: found.id ?? target.shape }, next)
    }

    this.touchSlide(entry.part)

    return result
  }

  setNotes(slide: number, value: TextInput): Changed {
    const entry = slides.slideAt(this.deck, slide)
    notes.writeNotes(this.deck, entry.part, value)
    return { changed: true }
  }

  deleteSlide(slide: number): Remaining {
    const entry = slides.slideAt(this.deck, slide)
    const pkg = this.package
    this.refuseLinked(entry, slide)
    const top = pkg.root(this.main)
    for (const sld of xml.descendants(top, P, 'sld')) {
      if (sld.getAttributeNS(R, 'id') === entry.rId) {
        xml.remove(sld)
      }
    }

    for (const ref of xml.descendants(top, P14, 'sldId')) {
      if (ref.getAttribute('id') === entry.id) {
        xml.remove(ref)
      }
    }

    xml.remove(entry.el)
    pkg.rels(this.main).remove(entry.rId)
    pkg.remove(entry.part)
    pkg.touch(this.main)
    this.slidesChanged = true

    return { slides: this.slideCount }
  }

  moveSlide(slide: number, opts: MoveOptions): Moved {
    slides.slideAt(this.deck, slide)
    const all = slides.entries(this.deck)
    const entry = all[slide]
    const to = isRecord(opts) ? opts.to : undefined
    if (typeof to !== 'number' || !Number.isInteger(to) || to < 0 || to >= all.length) {
      throw new DocsError(
        'SLIDE_NOT_FOUND',
        `to must be a position from 0 to ${all.length - 1}, got ${String(to)}`
      )
    }

    if (slide === to) {
      return { moved: false }
    }

    const list = entry.el.parentNode
    xml.remove(entry.el)
    const rest = all.filter((e) => e !== entry)
    list?.insertBefore(entry.el, rest[to]?.el ?? null)
    slides.placeInSection(this.deck, entry.id)
    this.package.touch(this.main)
    this.slidesChanged = true

    return { moved: true }
  }

  save(opts?: SaveOptions): Uint8Array {
    for (const part of this.touchedSlides) {
      if (this.package.has(part)) {
        this.package.dropUnreferenced(part, [REL.image, REL.chart, REL.hyperlink])
      }
    }

    this.touchedSlides = new Set()
    if (this.slidesChanged) {
      properties.refreshSummary(this.package, { Slides: this.slideCount })
      this.slidesChanged = false
    }

    properties.applySaveOptions(this.package, opts)

    return this.package.save()
  }

  private get deck(): slides.Deck {
    return { pkg: this.package, main: this.main }
  }

  private spTree(part: string): Element {
    const tree = xml.descendants(this.package.xml(part), P, 'spTree')[0]
    if (!tree) {
      throw new DocsError('BROKEN_PPTX', `${part} has no shape tree`)
    }

    return tree
  }

  private slideInfo(entry: slides.SlideEntry, index: number): SlideInfo {
    const doc = this.package.xml(entry.part)
    const layout = slides.layoutOf(this.package, entry.part)
    const info: SlideInfo = {
      index,
      layout: layout ? layoutName(this.package.xml(layout)) : null,
      shapes: shapesOf(doc).map((shape) => this.shapeInfo(entry.part, shape))
    }
    const noteText = notes.notesText(this.deck, entry.part)
    if (noteText) {
      info.notes = noteText
    }

    return info
  }

  private shapeInfo(part: string, shape: ShapeEntry): ShapeInfo {
    const out = describe(shape)
    const box = inherit.box({ pkg: this.package, slide: part, el: shape.el })
    if (box) {
      out.box = box
    }

    const chartPart = shape.kind === 'chart' ? this.chartPart(part, shape.el) : null
    if (chartPart) {
      out.chart = chart.readChart(this.package.xml(chartPart))
    }

    return out
  }

  private insertSlide(spec: NewSlide): void {
    const pkg = this.package
    const part = pkg.freeName('/ppt/slides/slide%d.xml')
    pkg.add(part, spec.doc, SLIDE_TYPE)
    pkg.rels(part).add(REL.slideLayout, spec.layout)
    const rId = pkg.rels(this.main).add(REL.slide, part)
    const list = slides.ensureSlideIdList(this.deck)
    const id = Math.max(255, ...spec.ids.map(Number)) + 1
    if (id >= 2147483648) {
      throw new DocsError('TOO_MANY_SLIDES', 'the deck has run out of slide ids')
    }

    const el = xml.builder(xml.owner(list), P)('sldId', { id })
    xml.setNsAttribute(el, { namespace: R, local: 'id' }, rId)
    list.insertBefore(el, spec.before)
    slides.placeInSection(this.deck, String(id))
    pkg.touch(this.main)
    this.slidesChanged = true
  }

  private refuseLinked(entry: slides.SlideEntry, slide: number): void {
    const pkg = this.package
    const noteParts = pkg.related(entry.part, REL.notesSlide)
    const linkers = pkg
      .names()
      .filter((name) => name !== this.main && name !== entry.part && !noteParts.includes(name))
      .filter((name) => pkg.rels(name).targeting(entry.part).length > 0)
    if (!linkers.length) {
      return
    }

    const all = slides.entries(this.deck)
    const where = linkers.map((name) => {
      const i = all.findIndex((s) => s.part === name)
      return i === -1 ? name : `slide ${i}`
    })

    throw new DocsError(
      'SLIDE_IS_LINKED',
      `slide ${slide} is the target of a link on ${where.join(', ')}; remove the link first`
    )
  }

  private touchSlide(part: string): void {
    this.touchedSlides.add(part)
    this.package.touch(part)
  }

  private locate(target: unknown): Located {
    if (!isRecord(target)) {
      throw new DocsError('SHAPE_NOT_FOUND', 'a shape is addressed as { slide, shape }')
    }

    const entry = slides.slideAt(this.deck, target.slide)
    const shapes = shapesOf(this.package.xml(entry.part))
    const found = findShape(shapes, target.shape)
    if (found) {
      return { entry, found }
    }

    throw shapeNotFound(target.slide, target.shape, shapes)
  }

  private write(at: Written, value: unknown): TextResult {
    const before = text.readParagraphs(at.txBody)
    if (sameParagraphs(before, withLevels(text.normalizeParagraphs(value), before))) {
      return { paragraphs: before.length, changed: false }
    }

    const removed = text.losses(at.txBody)
    const count = text.writeParagraphs(at.txBody, value)
    this.package.touch(at.part)
    const result: TextResult = { paragraphs: count, changed: true }
    if (Object.keys(removed).length) {
      result.removed = removed
    }

    return result
  }

  private estimate(part: string, el: Element, txBody: Element): fit.Estimate | null {
    const shape = { pkg: this.package, slide: part, el }
    const box = inherit.box(shape)
    if (!box) {
      return null
    }

    const paragraphs = xml.children(txBody, A, 'p')
    const sized = text.readParagraphs(txBody).map((p, i) => {
      const pEl = paragraphs[i]
      const run = pEl ? xml.descendants(pEl, A, 'rPr').find((r) => r.hasAttribute('sz')) : undefined
      const size = run ? Number(run.getAttribute('sz')) / 100 : inherit.fontSize(shape, p.level)
      return { text: p.text, level: p.level, size }
    })

    return fit.estimate(sized, box, inherit.insets(shape))
  }

  private chartOf(target: ShapeRef): string {
    const { entry, found } = this.locate(target)
    if (found.kind !== 'chart') {
      throw new DocsError('NOT_CHART', `shape ${found.id} is a ${found.kind}, not a chart`)
    }

    const part = this.chartPart(entry.part, found.el)
    if (!part) {
      throw new DocsError('BROKEN_PPTX', `chart ${found.id} points at a missing part`)
    }

    return part
  }

  private chartPart(slidePart: string, el: Element): string | null {
    const ref = xml.descendants(el, chart.C, 'chart')[0]
    const rel = ref ? this.package.rels(slidePart).get(ref.getAttributeNS(R, 'id')) : null
    return rel && !rel.external && rel.part && this.package.has(rel.part) ? rel.part : null
  }

  private box(opts: Placement): Box {
    checkPlacement(opts)
    const slide = this.slideSize
    const box = {
      x: opts.x ?? Math.round(slide.width * 0.07),
      y: opts.y ?? Math.round(slide.height * 0.25),
      width: opts.width ?? Math.round(slide.width * 0.86),
      height: opts.height ?? Math.round(slide.height * 0.65)
    }
    if (box.width === 0 || box.height === 0) {
      throw new DocsError('INVALID_BOX', 'the shape would have no visible size')
    }

    return box
  }

  private narrow(part: string, { shape, box }: Beside<SlideShape>): void {
    writeXfrm(shape.el, box)
    const txBody = xml.child(shape.el, P, 'txBody')
    const estimate = txBody ? this.estimate(part, shape.el, txBody) : null
    if (txBody && estimate?.overflow) {
      applyAutofit(txBody, estimate)
    }
  }

  private pictureArea(part: string, opts: Placement): Area<SlideShape> {
    if (opts.x !== undefined || opts.y !== undefined) {
      return { box: framed(this.slideSize), cleared: [], beside: null }
    }

    const shapes = shapesOf(this.package.xml(part)).flatMap((shape): SlideShape[] => {
      if (shape.group !== null || SKIPPED_PLACEHOLDERS.has(shape.placeholder ?? '')) {
        return []
      }

      const box = inherit.box({ pkg: this.package, slide: part, el: shape.el })

      return box ? [{ el: shape.el, box, role: roleOf(shape) }] : []
    })

    return pictureArea(this.slideSize, shapes)
  }

  private pictureBox(info: ImageInfo, opts: Placement, area: Box): Box {
    checkPlacement(opts)
    const ratio = info.height / info.width
    let { width, height } = opts
    if (width === undefined && height === undefined) {
      const natural = { width: info.width * EMU_PER_PIXEL, height: info.height * EMU_PER_PIXEL }
      const scale = Math.min(1, area.width / natural.width, area.height / natural.height)
      width = Math.round(natural.width * scale)
      height = Math.round(natural.height * scale)
    } else if (height === undefined) {
      height = Math.round((width ?? 0) * ratio)
    } else if (width === undefined) {
      width = Math.round(height / ratio)
    }

    if (!width || !height) {
      throw new DocsError('INVALID_BOX', 'the picture would have no visible size')
    }

    const x = opts.x ?? Math.max(0, area.x + Math.round((area.width - width) / 2))
    const y = opts.y ?? Math.max(0, area.y + Math.round((area.height - height) / 2))

    return { x, y, width, height }
  }
}
