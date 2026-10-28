import { Package, referencedIds } from '../opc/package.ts'
import * as xml from '../xml.ts'
import * as media from '../opc/media.ts'
import { probe, type ImageInfo } from '../image.ts'
import * as chart from '../drawingml/chart.ts'
import * as text from './text.ts'
import { Styles, type FormattedStyle } from './styles.ts'
import * as template from './template.ts'
import * as properties from '../opc/properties.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Document as XmlDocument, Element, Node } from '../xml.ts'
import type { OpenOptions } from '../opc/package.ts'
import type { SaveOptions } from '../opc/properties.ts'
import type { ChartInfo, ChartData, ChartSpec, DataResult, Changed } from '../drawingml/chart.ts'
import type { TextInput } from '../drawingml/text.ts'
import type { Alignment, Format, ParagraphSpec, Removed, Run } from './text.ts'
import * as units from '../units.ts'
import * as ids from './ids.ts'
import * as ranges from './ranges.ts'
import * as budget from './budget.ts'
import * as finder from './find.ts'
import * as outlines from './outline.ts'
import * as links from './links.ts'
import { readRuns, type RunView } from './runs.ts'
import type { BlockRef } from './ids.ts'
import type { Target } from './match.ts'
import type { Found } from './find.ts'
import type { Outline, OutlineOptions } from './outline.ts'
import type { TextQuery } from './match.ts'
import type { RunFormat } from './text.ts'
import type { FormatTextResult, LinkResult, ReplaceOptions, ReplaceResult } from './ranges.ts'
import * as host from './host.ts'
import * as lists from './lists.ts'
import * as markdownIn from './markdown-in.ts'
import * as markdownOut from './markdown-out.ts'
import * as comments from './comments.ts'
import * as footnotes from './footnotes.ts'
import * as page from './page.ts'
import * as toc from './toc.ts'
import * as revisions from './revisions.ts'
import * as templates from '../opc/templates.ts'
import type { CommentAdded, CommentDeleted, CommentResolved, ResolveOptions } from './comments.ts'
import type { CommentOptions } from './comments.ts'
import type { FootnoteAdded, FootnoteOptions } from './footnotes.ts'
import type { MarkdownOptions, MarkdownResult } from './markdown-out.ts'
import type { PageOptions, PageResult, PageSetup } from './page.ts'
import type { TocOptions } from './toc.ts'
import type { ChangeCounts } from './revisions.ts'
import type { DocumentConversion } from '../opc/templates.ts'

export type { Alignment, Format, ParagraphSpec, Removed, Run }
export type { CreateOptions } from './template.ts'
export type { ListKind } from './lists.ts'
export type { CommentAdded, CommentDeleted, CommentOptions, CommentResolved, ResolveOptions }
export type { FootnoteAdded, FootnoteOptions, MarkdownOptions, MarkdownResult }
export type { PageOptions, PageResult, PageSetup, TocOptions, ChangeCounts, DocumentConversion }
export type { Orientation, PageMargins, PageName, PageSize } from './page.ts'
export type TextSpec = string | ParagraphSpec
export type Paragraphs = TextSpec | TextSpec[]

export interface ReadOptions {
  from?: number
  count?: number
  maxChars?: number
  detail?: 'runs'
}

export interface ParagraphBlock {
  index: number
  id?: string
  type: 'paragraph'
  style: string
  text: string
  heading?: number
  list?: boolean
  level?: number
  footnotes?: string[]
  comments?: string[]
  images?: number
  charts?: ChartInfo[]
  pageBreak?: boolean
  runs?: RunView[]
}

export interface TableBlock {
  index: number
  id?: string
  type: 'table'
  rows: string[][]
}

export interface OtherBlock {
  index: number
  id?: string
  type: 'other'
  text: string
}

export type Block = ParagraphBlock | TableBlock | OtherBlock

export interface Note {
  id: string
  text: string
}

export interface Comment {
  id: string
  author: string | null
  text: string
}

export interface ReadResult {
  total: number
  blocks: Block[]
  next?: number
  headers: string[]
  footers: string[]
  footnotes: Note[]
  comments: Comment[]
}

export interface TextResult {
  changed: boolean
  removed?: Removed
}

export interface InsertResult {
  index: number
  count: number
  ids: string[]
}

export interface DeleteOptions {
  count?: number
}

export interface BlocksResult {
  blocks: number
}

export interface Section {
  index: number
  level: number
  title: string
  end: number
}

export interface SectionOptions {
  title?: TextSpec
  body?: Paragraphs
}

export interface SectionResult {
  index: number
  removed: number
  added: number
  ids: string[]
}

export interface TableOptions {
  rows: (string | number | null)[][]
  header?: boolean
  style?: string
}

export interface TableResult {
  index: number
  id: string
  rows: number
  columns: number
}

export interface CellRef {
  table: BlockRef
  row: number
  column: number
}

export interface RowRef {
  table: BlockRef
  row: number
}

export interface RowOptions {
  at?: number
}

export interface RowsResult {
  rows: number
}

export interface ImageOptions {
  width?: number
  height?: number
  description?: string
  name?: string
  align?: Alignment
}

export interface Size {
  width: number
  height: number
}

export interface Placed extends Size {
  index: number
  id: string
}

export interface DocChartSpec extends ChartSpec {
  width?: number
  height?: number
  name?: string
}

export interface ChartOptions {
  chart?: number
}

export interface HeaderOptions {
  type?: 'default' | 'first'
}

export interface TocResult extends InsertResult {
  levels: number
}

export interface ChangeOptions {
  in?: BlockRef
}

export interface HeaderResult {
  changed: boolean
  parts?: number
}

interface Anchor {
  anchor: Element | null
  index: number
  trailer: Element | null
}

interface Frame {
  id: number
  size: Size
  name: string
  description?: string
}

interface PictureFrame extends Frame {
  rId: string
}

interface Placement {
  blocks: Element[]
  at: number
}

interface HeadingAt {
  index: number
  level: number
}

interface HeaderPart {
  kind: 'header' | 'footer'
  type: string
}

const { w: W, r: R, wp: WP, a: A, pic: PIC } = xml.ns

const STRICT_REL = 'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument'
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
const REL = {
  officeDocument: REL_BASE + 'officeDocument',
  styles: REL_BASE + 'styles',
  header: REL_BASE + 'header',
  footer: REL_BASE + 'footer',
  chart: REL_BASE + 'chart',
  image: REL_BASE + 'image',
  hyperlink: REL_BASE + 'hyperlink',
  footnotes: REL_BASE + 'footnotes',
  endnotes: REL_BASE + 'endnotes',
  comments: REL_BASE + 'comments'
}

const DROPPABLE = [REL.image, REL.chart, REL.hyperlink, REL.header, REL.footer]
const DRAWING_PARTS = [REL.header, REL.footer, REL.footnotes, REL.endnotes, REL.comments]

const HEADER_TYPES = {
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
}

const SECTION_ORDER = [
  'headerReference',
  'footerReference',
  'footnotePr',
  'endnotePr',
  'type',
  'pgSz',
  'pgMar',
  'paperSrc',
  'pgBorders',
  'lnNumType',
  'pgNumType',
  'cols',
  'formProt',
  'vAlign',
  'noEndnote',
  'titlePg',
  'textDirection',
  'bidi',
  'rtlGutter',
  'docGrid',
  'printerSettings',
  'sectPrChange'
]

const MAIN_TYPES = new Set([
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
  'application/vnd.ms-word.document.macroEnabled.main+xml',
  'application/vnd.ms-word.template.macroEnabledTemplate.main+xml'
])

const BLOCKS = new Set(['p', 'tbl', 'sdt'])
const WRAPPERS = new Set(['sdt', 'customXml'])
const CELL_BLOCKS = new Set(['p', 'tbl', 'sdt', 'customXml'])
const WRAPPER_PROPERTIES = new Set(['sdtPr', 'sdtEndPr', 'customXmlPr'])
const UNWRAPPED = ['sdt', 'customXml', 'ins', 'moveTo']
const UNTRACKED = [
  'del',
  'moveFrom',
  'rPrChange',
  'pPrChange',
  'trPrChange',
  'tcPrChange',
  'tblPrExChange',
  'sectPrChange',
  'numberingChange',
  'cellIns',
  'cellDel',
  'cellMerge',
  'moveFromRangeStart',
  'moveFromRangeEnd',
  'moveToRangeStart',
  'moveToRangeEnd',
  'bookmarkStart',
  'bookmarkEnd',
  'commentRangeStart',
  'commentRangeEnd',
  'permStart',
  'permEnd'
]
const NOTE_REFERENCES = ['commentReference', 'footnoteReference', 'endnoteReference']
const MARKS = new Set(['trPr', 'rPr'])
const NOTE_TYPES = ['separator', 'continuationSeparator', 'continuationNotice']
const PIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
const EMU_PER_TWIP = 635
const EMU_PER_PIXEL = 9525

function isBlock(el: Element): boolean {
  return BLOCKS.has(el.localName ?? '')
}

function blocksOf(parent: Element): Element[] {
  return xml.children(parent, W).filter(isBlock)
}

function textOf(parent: Element): string {
  return blocksOf(parent).map(text.blockText).join('\n').trim()
}

function withRemoved(result: TextResult, removed: Removed): TextResult {
  if (Object.keys(removed).length) {
    result.removed = removed
  }

  return result
}

function validateRows(rows: unknown): (string | number | null)[][] {
  if (!Array.isArray(rows) || rows.length === 0 || !rows.every(Array.isArray)) {
    throw new DocsError(
      'INVALID_TABLE',
      'rows must be a non-empty array of rows, each an array of cells'
    )
  }

  const columns = Math.max(...rows.map((r) => r.length))
  if (columns === 0) {
    throw new DocsError('INVALID_TABLE', 'the table has no columns')
  }

  if (rows.length > 5000 || columns > 63) {
    throw new DocsError('INVALID_TABLE', 'a table holds at most 5000 rows and 63 columns')
  }

  return rows.map((row: unknown[], r) =>
    row.map((value, c) => {
      if (value === null || typeof value === 'string') return value
      if (typeof value === 'number' && Number.isFinite(value)) return value
      throw new DocsError('INVALID_TABLE', `cell ${r},${c} must be a string, a number or null`)
    })
  )
}

function cellText(value: unknown): string {
  return value === null || value === undefined ? '' : String(value)
}

function borders(doc: XmlDocument): Element {
  const el = text.wEl(doc, 'tblBorders')
  for (const side of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) {
    el.appendChild(text.wEl(doc, side, { val: 'single', sz: 4, space: 0, color: 'auto' }))
  }

  return el
}

function inlineDrawing(graphic: Element, frame: Frame): Element {
  const doc = xml.owner(graphic)
  const wp = xml.builder(doc, WP)
  const a = xml.builder(doc, A)
  const drawing = text.wEl(doc, 'drawing')
  const inline = xml.append(drawing, wp('inline', { distT: 0, distB: 0, distL: 0, distR: 0 }))
  inline.appendChild(wp('extent', { cx: frame.size.width, cy: frame.size.height }))
  inline.appendChild(wp('effectExtent', { l: 0, t: 0, r: 0, b: 0 }))
  inline.appendChild(wp('docPr', { id: frame.id, name: frame.name, descr: frame.description }))
  xml
    .append(inline, wp('cNvGraphicFramePr'))
    .appendChild(a('graphicFrameLocks', { noChangeAspect: 1 }))
  inline.appendChild(graphic)

  return drawing
}

function pictureGraphic(doc: XmlDocument, frame: PictureFrame): Element {
  const a = xml.builder(doc, A)
  const pic = xml.builder(doc, PIC)
  const graphic = a('graphic')
  const data = xml.append(graphic, a('graphicData', { uri: PIC_URI }))
  const picture = xml.append(data, pic('pic'))
  const nv = xml.append(picture, pic('nvPicPr'))
  nv.appendChild(pic('cNvPr', { id: frame.id, name: frame.name, descr: frame.description }))
  nv.appendChild(pic('cNvPicPr'))
  const fill = xml.append(picture, pic('blipFill'))
  const blip = xml.append(fill, a('blip'))
  xml.setNsAttribute(blip, { namespace: R, local: 'embed' }, frame.rId)
  xml.append(fill, a('stretch')).appendChild(a('fillRect'))
  const spPr = xml.append(picture, pic('spPr'))
  const xfrm = xml.append(spPr, a('xfrm'))
  xfrm.appendChild(a('off', { x: 0, y: 0 }))
  xfrm.appendChild(a('ext', { cx: frame.size.width, cy: frame.size.height }))
  xml.append(spPr, a('prstGeom', { prst: 'rect' })).appendChild(a('avLst'))
  return graphic
}

function chartGraphic(doc: XmlDocument, rId: string): Element {
  const a = xml.builder(doc, A)
  const graphic = a('graphic')
  const data = xml.append(graphic, a('graphicData', { uri: chart.GRAPHIC_URI }))
  const ref = xml.append(data, xml.named(doc, 'c:chart', chart.C))
  ref.setAttributeNS(R, 'r:id', rId)
  return graphic
}

function isEmptyParagraph(p: Element): boolean {
  return xml.children(p, W).every((n) => n.localName === 'pPr')
}

function contentOf(wrapper: Element): Element {
  return xml.child(wrapper, W, 'sdtContent') ?? wrapper
}

function unwrapped(parent: Element, local: string): Element[] {
  const out: Element[] = []
  for (const el of xml.children(parent, W)) {
    if (el.localName === local) {
      out.push(el)
    } else if (WRAPPERS.has(el.localName ?? '')) {
      out.push(...unwrapped(contentOf(el), local))
    }
  }

  return out
}

function firstParagraph(parent: Element): Element | null {
  for (const el of xml.children(parent, W)) {
    if (el.localName === 'p') {
      return el
    }

    const found = WRAPPERS.has(el.localName ?? '') ? firstParagraph(contentOf(el)) : null
    if (found) {
      return found
    }
  }

  return null
}

function clearCell(tc: Element): Element {
  const keep = firstParagraph(tc)
  if (!keep) {
    for (const el of xml.children(tc, W).filter((n) => CELL_BLOCKS.has(n.localName ?? ''))) {
      xml.remove(el)
    }

    return xml.append(tc, text.wEl(xml.owner(tc), 'p'))
  }

  for (let node: Node = keep; node !== tc && node.parentNode; node = node.parentNode) {
    for (const el of xml.children(node.parentNode, W)) {
      if (el !== node && CELL_BLOCKS.has(el.localName ?? '')) {
        xml.remove(el)
      }
    }
  }

  clearPlaceholders(keep)

  return keep
}

function clearPlaceholders(from: Element): void {
  for (let node = from.parentNode; node && !xml.isElement(node, W, 'tbl'); node = node.parentNode) {
    const sdtPr = xml.isElement(node, W, 'sdt') ? xml.child(node, W, 'sdtPr') : null
    for (const el of sdtPr ? xml.children(sdtPr, W, 'showingPlcHdr') : []) {
      xml.remove(el)
    }
  }
}

function unwrap(el: Element): void {
  const parent = el.parentNode
  if (!parent) {
    return
  }

  for (const child of xml.children(contentOf(el))) {
    if (!WRAPPER_PROPERTIES.has(child.localName ?? '')) {
      parent.insertBefore(child, el)
    }
  }

  xml.remove(el)
}

function dropHeaderMark(tr: Element): void {
  const trPr = xml.child(tr, W, 'trPr')
  for (const h of trPr ? xml.children(trPr, W, 'tblHeader') : []) {
    xml.remove(h)
  }
}

interface RowPlace {
  tbl: Element
  rows: Element[]
  at: number
}

function insertRow(tr: Element, { tbl, rows, at }: RowPlace): void {
  const before = rows[at - 1]
  if (before?.parentNode) {
    before.parentNode.insertBefore(tr, before.nextSibling)
  } else {
    ;(rows[0]?.parentNode ?? tbl).insertBefore(tr, rows[0] ?? null)
  }
}

function plainRow(tr: Element): Element {
  for (let el = firstOf(tr, UNWRAPPED); el; el = firstOf(tr, UNWRAPPED)) {
    if (MARKS.has(el.parentNode?.localName ?? '')) {
      xml.remove(el)
    } else {
      unwrap(el)
    }
  }

  for (const local of UNTRACKED) {
    for (const el of xml.descendants(tr, W, local)) {
      xml.remove(el)
    }
  }

  for (const r of xml.descendants(tr, W, 'r')) {
    if (NOTE_REFERENCES.some((ref) => xml.child(r, W, ref))) {
      xml.remove(r)
    }
  }

  return tr
}

function firstOf(el: Element, locals: readonly string[]): Element | null {
  for (const local of locals) {
    const found = xml.descendants(el, W, local)[0]
    if (found) {
      return found
    }
  }

  return null
}

function removeEmptyWrappers(from: Node | null, stop: Element): void {
  let node = from
  while (node && node !== stop && xml.isElement(node, W)) {
    const wrapper = node.localName === 'sdtContent' ? node.parentNode : node
    if (!xml.isElement(wrapper, W) || !WRAPPERS.has(wrapper.localName ?? '')) {
      return
    }

    const content = xml.children(contentOf(wrapper))
    if (content.some((el) => !WRAPPER_PROPERTIES.has(el.localName ?? ''))) {
      return
    }

    node = wrapper.parentNode
    xml.remove(wrapper)
  }
}

function checkPositive(values: Record<string, unknown>): void {
  for (const [key, value] of Object.entries(values)) {
    if (
      value !== undefined &&
      !(typeof value === 'number' && Number.isInteger(value) && value > 0)
    ) {
      throw new DocsError(
        'INVALID_BOX',
        `${key} must be a positive whole number of EMU, got ${String(value)}`
      )
    }
  }
}

function mainPart(pkg: Package): string {
  const main = pkg.related('/', REL.officeDocument)[0]
  if (main && MAIN_TYPES.has(pkg.contentType(main) ?? '')) {
    return main
  }

  if (pkg.rels('/').byType(STRICT_REL).length) {
    throw new DocsError(
      'UNSUPPORTED_DOCX',
      'the document uses Strict Open XML, which is not supported'
    )
  }

  throw new DocsError('NOT_DOCX', 'the file is not a Word document')
}

export class Document {
  package: Package
  main: string
  styles: Styles
  private orphans: Set<string>

  constructor(pkg: Package) {
    this.package = pkg
    this.main = mainPart(pkg)
    const top = pkg.root(this.main)
    if (top.namespaceURI !== W || top.localName !== 'document') {
      throw new DocsError(
        'UNSUPPORTED_DOCX',
        'the document uses Strict Open XML, which is not supported'
      )
    }

    this.styles = new Styles(pkg, pkg.related(this.main, REL.styles)[0] ?? null)
    this.orphans = new Set(this.unreferenced())
  }

  static open(bytes: Uint8Array | ArrayBuffer, opts?: OpenOptions): Document {
    return new Document(Package.open(bytes, opts))
  }

  static create(opts: template.CreateOptions = {}): Document {
    const pkg = Package.create()
    template.build(pkg, opts)
    return new Document(pkg)
  }

  get changed(): boolean {
    return this.package.changed
  }

  get blockCount(): number {
    return this.blocks().length
  }

  read(opts: ReadOptions = {}): ReadResult {
    const all = this.blocks()
    const from = opts.from ?? 0
    const count = opts.count ?? Math.max(all.length, 1)
    if (!Number.isInteger(from) || from < 0 || (from >= all.length && all.length > 0)) {
      throw new DocsError(
        'INVALID_OPTIONS',
        `from must be a block index from 0 to ${Math.max(0, all.length - 1)}, got ${String(opts.from)}`
      )
    }

    if (!Number.isInteger(count) || count < 1) {
      throw new DocsError('INVALID_OPTIONS', `count must be at least 1, got ${String(opts.count)}`)
    }

    const runs = detailRuns(opts.detail)
    const window = {
      from,
      end: Math.min(all.length, from + count),
      total: all.length,
      maxChars: budget.validateMaxChars(opts.maxChars)
    }
    const page = budget.take(window, (i) => this.readBlock(all[i], i, runs))
    const result: ReadResult = {
      total: all.length,
      blocks: page.items,
      headers: this.partTexts(REL.header),
      footers: this.partTexts(REL.footer),
      footnotes: this.notes(),
      comments: this.comments()
    }
    if (page.next !== undefined) {
      result.next = page.next
    }

    return result
  }

  setHeader(value: Paragraphs | null, opts: HeaderOptions = {}): HeaderResult {
    return this.setHeaderFooter({ kind: 'header', type: headerType(opts) }, value)
  }

  setFooter(value: Paragraphs | null, opts: HeaderOptions = {}): HeaderResult {
    return this.setHeaderFooter({ kind: 'footer', type: headerType(opts) }, value)
  }

  paragraphStyles(): string[] {
    return this.styles.names('paragraph')
  }

  setText(ref: BlockRef, input: TextSpec): TextResult {
    const index = this.indexOf(ref)
    const p = this.paragraph(index)
    const before = text.paragraphText(p)
    const runs = text.normalizeRuns(input)
    const plain = runs.every((run) => Object.keys(run).length === 1)
    if (plain && runs.map((run) => run.text).join('') === before) {
      return { changed: false }
    }

    text.assertUntracked(p, `block ${index}`)
    const removed = text.replaceContent(p, input)
    this.touch()

    return withRemoved({ changed: true }, removed)
  }

  insertParagraphs(where: BlockRef, specs: Paragraphs): InsertResult {
    const blocks = this.blocks()
    const at = this.at(where, blocks)
    const list = this.specs(specs)
    const before = this.templateBefore(blocks, at)
    const { anchor, index, trailer } = this.anchor(blocks, at)
    const body = this.body()
    const blank = this.blankOnly(blocks)
    const built = list.map((spec) => this.build(spec, before))
    const added = this.stamp(built)
    for (const p of built) {
      body.insertBefore(p, anchor)
    }

    if (blank) {
      xml.remove(blank)
    }

    if (trailer) {
      xml.remove(trailer)
    }

    this.touch()
    const shift = blank && index > 0 ? 1 : 0

    return { index: index - shift, count: list.length, ids: added }
  }

  appendParagraphs(specs: Paragraphs): InsertResult {
    return this.insertParagraphs(this.blockCount, specs)
  }

  deleteBlocks(ref: BlockRef, opts: DeleteOptions = {}): BlocksResult {
    const count = opts.count ?? 1
    const blocks = this.blocks()
    const index = ids.resolve(blocks, ref)
    this.block(index)
    if (!Number.isInteger(count) || count < 1 || index + count > blocks.length) {
      throw new DocsError(
        'BLOCK_NOT_FOUND',
        `count must be from 1 to ${blocks.length - index} when deleting from block ${index}`
      )
    }

    for (const el of blocks.slice(index, index + count)) {
      xml.remove(el)
    }

    this.ensureTrailingParagraph()
    this.touch()

    return { blocks: this.blockCount }
  }

  sections(): Section[] {
    const blocks = this.blocks()
    const out: Section[] = []
    blocks.forEach((el, index) => {
      const level = this.headingLevel(el)
      if (level === null || level < 1) {
        return
      }

      out.push({
        index,
        level,
        title: text.paragraphText(el),
        end: this.sectionEnd(blocks, { index, level })
      })
    })

    return out
  }

  replaceSection(ref: BlockRef, opts: SectionOptions = {}): SectionResult {
    const blocks = this.blocks()
    const index = ids.resolve(blocks, ref)
    const level = this.heading(blocks, index)
    const end = this.sectionEnd(blocks, { index, level })
    const body = blocks.slice(index + 1, end)
    const list = opts.body === undefined ? null : this.specs(opts.body)
    const built: Element[] = []
    if (opts.title !== undefined) {
      text.assertUntracked(blocks[index], `heading ${index}`)
      text.replaceContent(blocks[index], opts.title)
    }

    if (list) {
      const model = body.find(
        (el) => el.localName === 'p' && this.headingLevel(el) === null && !this.isList(el)
      )
      const anchor = blocks[end] ?? this.sectPr()
      built.push(...list.map((spec) => this.build(spec, model ?? null)))
      for (const p of built) {
        this.body().insertBefore(p, anchor)
      }

      for (const el of body) {
        xml.remove(el)
      }
    }

    this.touch()

    return {
      index,
      removed: list ? body.length : 0,
      added: list ? list.length : 0,
      ids: this.stamp(built)
    }
  }

  deleteSection(ref: BlockRef): BlocksResult {
    const blocks = this.blocks()
    const index = ids.resolve(blocks, ref)
    const level = this.heading(blocks, index)
    const end = this.sectionEnd(blocks, { index, level })
    return this.deleteBlocks(index, { count: end - index })
  }

  formatParagraph(ref: BlockRef, input: Format): Changed {
    const p = this.paragraph(this.indexOf(ref))
    const { align, ...run } = text.validateFormat(input)
    if (align) {
      text.setAlign(p, align)
    }

    if (Object.keys(run).length) {
      for (const r of xml.descendants(p, W, 'r')) {
        if (!xml.child(r, W, 't')) continue
        let rPr = xml.child(r, W, 'rPr')
        if (!rPr) {
          rPr = text.wEl(xml.owner(r), 'rPr')
          r.insertBefore(rPr, r.firstChild)
        }

        text.applyRunFormat(rPr, run)
      }
    }

    this.touch()

    return { changed: true }
  }

  format(ref: BlockRef, input: Format): Changed {
    return this.formatParagraph(ref, input)
  }

  formatStyle(name: string, input: Format): FormattedStyle {
    if (!this.styles.part) {
      throw new DocsError('STYLE_NOT_FOUND', 'the document has no styles part')
    }

    return this.styles.format(name, input)
  }

  addTable(where: BlockRef, opts: TableOptions): TableResult {
    const blocks = this.blocks()
    const at = this.at(where, blocks)
    const rows = validateRows(isRecord(opts) ? opts.rows : undefined)
    const columns = Math.max(...rows.map((r) => r.length))
    const header = opts.header !== false
    const tbl = this.tableElement({ rows, columns, header, styleId: this.tableStyleId(opts.style) })
    const [id] = this.stamp([tbl])
    const { anchor, index } = this.anchor(blocks, at)
    this.body().insertBefore(tbl, anchor)
    this.ensureTrailingParagraph()
    this.touch()
    return { index, id, rows: rows.length, columns }
  }

  setCell(cell: CellRef, input: TextSpec): TextResult {
    const tc = this.cell(cell)
    text.assertUntracked(tc, `cell ${cell.row},${cell.column}`)
    const removed = text.replaceContent(clearCell(tc), input)
    this.touch()
    return withRemoved({ changed: true }, removed)
  }

  addRow(table: BlockRef, values: (string | number | null)[], opts: RowOptions = {}): RowsResult {
    const tbl = this.table(table)
    const rows = unwrapped(tbl, 'tr')
    const at = opts.at ?? rows.length
    if (!Number.isInteger(at) || at < 0 || at > rows.length) {
      throw new DocsError(
        'ROW_NOT_FOUND',
        `at must be a row position from 0 to ${rows.length}, got ${String(opts.at)}`
      )
    }

    if (!Array.isArray(values)) {
      throw new DocsError('INVALID_TABLE', 'values must be an array of cell texts')
    }

    const model = at > 0 && rows[at - 1] ? rows[at - 1] : rows[rows.length > 1 ? 1 : 0]
    const tr = plainRow(xml.clone(model))
    dropHeaderMark(tr)
    const cells = xml.children(tr, W, 'tc')
    if (values.length > cells.length) {
      throw new DocsError(
        'INVALID_TABLE',
        `the row has ${cells.length} cells but ${values.length} values were given`
      )
    }

    cells.forEach((tc, c) => fillCell(tc, cellText(values[c])))
    this.stamp([tr])
    insertRow(tr, { tbl, rows, at })
    this.touch()

    return { rows: rows.length + 1 }
  }

  deleteRow(ref: RowRef): RowsResult {
    const tbl = this.table(ref.table)
    const rows = unwrapped(tbl, 'tr')
    const tr = rows[ref.row]
    if (!tr) {
      throw new DocsError(
        'ROW_NOT_FOUND',
        `row ${ref.row} does not exist; the table has ${rows.length} (0 to ${rows.length - 1})`
      )
    }

    if (rows.length === 1) {
      throw new DocsError(
        'LAST_ROW',
        "this is the table's only row; delete the table with deleteBlocks instead"
      )
    }

    if (xml.descendants(tr, W, 'vMerge').length) {
      throw new DocsError(
        'ROW_MERGED',
        `row ${ref.row} is part of a vertically merged cell; unmerge it in Word first`
      )
    }

    const parent = tr.parentNode
    xml.remove(tr)
    removeEmptyWrappers(parent, tbl)
    this.touch()

    return { rows: rows.length - 1 }
  }

  addImage(
    where: BlockRef,
    bytes: Uint8Array,
    input: units.WithLengths<ImageOptions, units.SizeKey> = {}
  ): Placed {
    const opts = units.size(input)
    const blocks = this.blocks()
    const at = this.at(where, blocks)
    const size = this.imageSize(probe(bytes), opts)
    const { rId } = media.addImage(this.package, bytes, { source: this.main, dir: 'media' })
    const doc = this.doc()
    const id = this.nextDocPrId()
    const p = text.wEl(doc, 'p')
    if (opts.align) {
      const { align } = text.validateFormat({ align: opts.align })
      if (align) {
        text.setAlign(p, align)
      }
    }

    const frame = { id, size, name: opts.name || `Picture ${id}`, description: opts.description }
    xml
      .append(p, text.wEl(doc, 'r'))
      .appendChild(inlineDrawing(pictureGraphic(doc, { ...frame, rId }), frame))

    const [placed] = this.stamp([p])

    return { index: this.place(p, { blocks, at }), id: placed, ...size }
  }

  addChart(where: BlockRef, input: units.WithLengths<DocChartSpec, units.SizeKey>): Placed {
    const spec = units.size(input)
    const blocks = this.blocks()
    const at = this.at(where, blocks)
    const width = spec.width ?? this.textWidth() * EMU_PER_TWIP
    const height = spec.height ?? Math.round(width * 0.6)
    checkPositive({ width, height })
    const part = this.package.freeName('/word/charts/chart%d.xml')
    chart.createChart({ pkg: this.package, part }, spec)
    const rId = this.package.rels(this.main).add(REL.chart, part)
    const doc = this.doc()
    const id = this.nextDocPrId()
    const p = text.wEl(doc, 'p')
    const frame = { id, size: { width, height }, name: spec.name || `Chart ${id}` }
    xml.append(p, text.wEl(doc, 'r')).appendChild(inlineDrawing(chartGraphic(doc, rId), frame))
    const [placed] = this.stamp([p])
    return { index: this.place(p, { blocks, at }), id: placed, width, height }
  }

  setChartData(ref: BlockRef, data: ChartData, opts: ChartOptions = {}): DataResult {
    return chart.setChartData(
      { pkg: this.package, part: this.chartPart(this.indexOf(ref), opts.chart ?? 0) },
      data
    )
  }

  setChartTitle(ref: BlockRef, title: TextInput | null, opts: ChartOptions = {}): Changed {
    if (!(title === null || typeof title === 'string' || Array.isArray(title))) {
      throw new DocsError('INVALID_TEXT', 'title must be text, or null to remove it')
    }

    return chart.setChartTitle(
      { pkg: this.package, part: this.chartPart(this.indexOf(ref), opts.chart ?? 0) },
      title
    )
  }

  find(query: finder.FindQuery): Found[] {
    return finder.find(this.blocks(), (el) => this.summary(el), query)
  }

  outline(opts: OutlineOptions = {}): Outline {
    return outlines.outline(this.blocks(), this.sections(), opts)
  }

  replaceText(opts: ReplaceOptions): ReplaceResult {
    const result = ranges.replaceText(this.targets(opts.in), opts)
    this.touchIf(result.replaced)
    return result
  }

  formatText(ref: BlockRef, query: TextQuery, format: RunFormat): FormatTextResult {
    const result = ranges.formatText(this.targets(ref), query, format)
    this.touchIf(result.formatted)
    return result
  }

  link(ref: BlockRef, query: TextQuery, url: string): LinkResult {
    const target = links.validateUrl(url)
    const rel = (): string => links.addHyperlink(this.package.rels(this.main), target)
    const style = this.styles.resolve('Hyperlink', 'character')?.id ?? null
    const result = ranges.linkText(this.targets(ref), query, { rel, style })
    this.touchIf(result.linked)
    return result
  }

  appendMarkdown(markdown: string): InsertResult {
    return this.insertMarkdown(this.blockCount, markdown)
  }

  insertMarkdown(where: BlockRef, markdown: string): InsertResult {
    const at = this.at(where, this.blocks())
    return this.insertElements(at, markdownIn.build(this, markdown))
  }

  markdown(opts: MarkdownOptions = {}): MarkdownResult {
    return markdownOut.write(this, opts)
  }

  addComment(block: BlockRef, opts: CommentOptions): CommentAdded {
    return comments.add(this, host.paragraph(this, this.indexOf(block)), opts)
  }

  deleteComment(id: string | number): CommentDeleted {
    return comments.remove(this, id)
  }

  resolveComment(id: string | number, opts: ResolveOptions = {}): CommentResolved {
    return comments.resolve(this, id, opts)
  }

  addFootnote(block: BlockRef, opts: FootnoteOptions): FootnoteAdded {
    return footnotes.add(this, host.paragraph(this, this.indexOf(block)), opts)
  }

  setPage(opts: PageOptions): PageResult {
    return page.set(this, opts)
  }

  page(): PageSetup {
    return page.read(this)
  }

  addTableOfContents(where: BlockRef, opts: TocOptions = {}): TocResult {
    const at = this.at(where, this.blocks())
    const built = toc.build(this, opts)
    return { ...this.insertElements(at, built), levels: opts.levels ?? 3 }
  }

  acceptChanges(opts: ChangeOptions = {}): ChangeCounts {
    return this.resolveChanges(opts, 'accept')
  }

  rejectChanges(opts: ChangeOptions = {}): ChangeCounts {
    return this.resolveChanges(opts, 'reject')
  }

  get template(): boolean {
    return templates.isTemplate(this.package, this.main)
  }

  toDocument(): DocumentConversion {
    return templates.toDocument(this.package, this.main)
  }

  save(opts?: SaveOptions): Uint8Array {
    if (this.package.has(this.main)) {
      this.dropOrphaned()
    }

    properties.applySaveOptions(this.package, opts)

    return this.package.save()
  }

  private doc(): XmlDocument {
    return this.package.xml(this.main)
  }

  private unreferenced(): string[] {
    const used = referencedIds(this.package.root(this.main))
    return this.package
      .rels(this.main)
      .all()
      .filter((rel) => DROPPABLE.includes(rel.type) && !used.has(rel.id))
      .map((rel) => rel.id)
  }

  private dropOrphaned(): void {
    const rels = this.package.rels(this.main)
    const dropped = this.unreferenced().filter((id) => !this.orphans.has(id))
    for (const id of dropped) {
      rels.remove(id)
    }

    if (dropped.length) {
      this.package.prune()
    }
  }

  private body(): Element {
    const body = xml.child(this.package.root(this.main), W, 'body')
    if (!body) {
      throw new DocsError('BROKEN_DOCX', 'the document has no body')
    }

    return body
  }

  private sectPr(): Element | null {
    return xml.child(this.body(), W, 'sectPr')
  }

  private blocks(): Element[] {
    return blocksOf(this.body())
  }

  private touch(): void {
    this.package.touch(this.main)
  }

  private block(index: number): Element {
    const blocks = this.blocks()
    const found = Number.isInteger(index) && index >= 0 ? blocks[index] : undefined
    if (!found) {
      const range = blocks.length ? ` (0 to ${blocks.length - 1})` : ''
      throw new DocsError(
        'BLOCK_NOT_FOUND',
        `block ${index} does not exist; the document has ${blocks.length} blocks${range}`
      )
    }

    return found
  }

  private paragraph(index: number): Element {
    const el = this.block(index)
    if (el.localName === 'p') {
      return el
    }

    const kind =
      el.localName === 'tbl' ? 'a table; use setCell' : 'a content control and cannot be edited'

    throw new DocsError('NOT_PARAGRAPH', `block ${index} is ${kind}`)
  }

  private table(ref: BlockRef): Element {
    const index = this.indexOf(ref)
    const el = this.block(index)
    if (el.localName !== 'tbl') {
      throw new DocsError('NOT_TABLE', `block ${index} is not a table`)
    }

    return el
  }

  private cell(ref: CellRef): Element {
    const { row, column } = ref
    const rows = unwrapped(this.table(ref.table), 'tr')
    const tr = Number.isInteger(row) && row >= 0 ? rows[row] : undefined
    if (!tr) {
      throw new DocsError(
        'CELL_NOT_FOUND',
        `row ${row} does not exist; the table has ${rows.length} (0 to ${rows.length - 1})`
      )
    }

    const cells = unwrapped(tr, 'tc')
    const tc = Number.isInteger(column) && column >= 0 ? cells[column] : undefined
    if (!tc) {
      throw new DocsError(
        'CELL_NOT_FOUND',
        `column ${column} does not exist; the row has ${cells.length} (0 to ${cells.length - 1})`
      )
    }

    return checkMerge(tc, ref)
  }

  private anchor(blocks: Element[], at: number): Anchor {
    if (at < blocks.length) {
      return { anchor: blocks[at], index: at, trailer: null }
    }

    const last = blocks[blocks.length - 1]
    const before = blocks[blocks.length - 2]
    const trailer =
      !!last &&
      !!before &&
      before.localName === 'tbl' &&
      last.localName === 'p' &&
      isEmptyParagraph(last)
    if (trailer) {
      return { anchor: last, index: at - 1, trailer: last }
    }

    return { anchor: this.sectPr(), index: at, trailer: null }
  }

  private place(p: Element, where: Placement): number {
    const { anchor, index, trailer } = this.anchor(where.blocks, where.at)
    this.body().insertBefore(p, anchor)
    if (trailer) {
      xml.remove(trailer)
    }

    this.touch()

    return index
  }

  private position(at: number, length: number): void {
    if (!Number.isInteger(at) || at < 0 || at > length) {
      throw new DocsError('BLOCK_NOT_FOUND', `at must be a position from 0 to ${length}, got ${at}`)
    }
  }

  private specs(specs: unknown): unknown[] {
    const list: unknown[] = Array.isArray(specs) ? specs : [specs]
    if (list.length === 0) {
      throw new DocsError('INVALID_TEXT', 'give at least one paragraph')
    }

    const styled = list.map((spec) => {
      if (isRecord(spec) && spec.style !== undefined) {
        return { ...spec, styleId: this.styles.require(spec.style).id }
      }

      return spec
    })

    return lists.assign(this, styled)
  }

  private build(spec: unknown, model: Element | null): Element {
    const p = text.buildParagraph(this.doc(), spec, model)
    if (isRecord(spec) && spec.list !== undefined) {
      lists.apply(p, spec)
    } else if (isRecord(spec) && spec.level !== undefined) {
      this.setLevel(p, spec.level)
    }

    return p
  }

  private setLevel(p: Element, level: unknown): void {
    if (!(typeof level === 'number' && Number.isInteger(level) && level >= 0 && level <= 8)) {
      throw new DocsError(
        'INVALID_TEXT',
        `level must be an integer from 0 to 8, got ${String(level)}`
      )
    }

    const pPr = text.ensure(p, 'pPr')
    const own = xml.child(pPr, W, 'numPr')
    const numId =
      text.wGet(own ? xml.child(own, W, 'numId') : null, 'val') ||
      this.styles.numIdOf(this.styleIdOf(p))
    if (!numId || numId === '0') {
      throw new DocsError(
        'INVALID_TEXT',
        'level applies to list paragraphs; give a list style such as "List Bullet"'
      )
    }

    if (own) {
      xml.remove(own)
    }

    const doc = xml.owner(p)
    const numPr = xml.insertOrdered(pPr, text.wEl(doc, 'numPr'), text.PARAGRAPH_ORDER)
    numPr.appendChild(text.wEl(doc, 'ilvl', { val: level }))
    numPr.appendChild(text.wEl(doc, 'numId', { val: numId }))
  }

  private sectionList(): Element[] {
    return xml
      .descendants(this.body(), W, 'sectPr')
      .filter((el) => !xml.isElement(el.parentNode, W, 'sectPrChange'))
  }

  private setHeaderFooter(target: HeaderPart, value: Paragraphs | null): HeaderResult {
    const sections = this.sectionList()
    if (sections.length === 0) {
      throw new DocsError('BROKEN_DOCX', 'the document has no section properties')
    }

    const refs = sections.flatMap((sectPr) =>
      xml
        .children(sectPr, W, `${target.kind}Reference`)
        .filter((r) => (text.wGet(r, 'type') || 'default') === target.type)
    )
    if (value === null) {
      for (const ref of refs) {
        xml.remove(ref)
      }

      this.touch()

      return { changed: refs.length > 0 }
    }

    const paragraphs = headerParagraphs(value, target.kind)

    const parts = this.headerParts(refs)
    if (parts.length === 0) {
      parts.push(this.createHeaderPart(target, sections[0]))
    }

    for (const part of parts) {
      this.fillHeaderPart(part, paragraphs)
    }

    return { changed: true, parts: parts.length }
  }

  private headerParts(refs: Element[]): string[] {
    const rels = this.package.rels(this.main)
    const parts = new Set<string>()
    for (const ref of refs) {
      const rel = rels.get(ref.getAttributeNS(R, 'id'))
      if (rel && !rel.external && rel.part && this.package.has(rel.part)) {
        parts.add(rel.part)
      }
    }

    return [...parts]
  }

  private createHeaderPart(target: HeaderPart, first: Element): string {
    const pkg = this.package
    const part = pkg.freeName(`/word/${target.kind}%d.xml`)
    const style = this.styles.resolve(target.kind === 'header' ? 'Header' : 'Footer')
    const rootName = target.kind === 'header' ? 'hdr' : 'ftr'
    const doc = xml.parse(`<w:${rootName} xmlns:w="${W}" xmlns:r="${R}"><w:p/></w:${rootName}>`)
    const p = xml.child(xml.root(doc), W, 'p')
    if (style && p) {
      text.ensure(p, 'pPr').appendChild(text.wEl(doc, 'pStyle', { val: style.id }))
    }

    pkg.add(part, doc, HEADER_TYPES[target.kind])
    const rId = pkg.rels(this.main).add(target.kind === 'header' ? REL.header : REL.footer, part)
    const main = this.doc()
    const ref = xml.insertOrdered(
      first,
      text.wEl(main, `${target.kind}Reference`, { type: target.type }),
      SECTION_ORDER
    )
    xml.setNsAttribute(ref, { namespace: R, local: 'id' }, rId)
    if (target.type === 'first' && !xml.child(first, W, 'titlePg')) {
      xml.insertOrdered(first, text.wEl(main, 'titlePg'), SECTION_ORDER)
    }

    this.touch()

    return part
  }

  private fillHeaderPart(part: string, paragraphs: unknown[]): void {
    const doc = this.package.xml(part)
    const top = xml.root(doc)
    const blocks = blocksOf(top)
    const model = blocks.find((el) => el.localName === 'p') ?? null
    for (const spec of paragraphs) {
      top.insertBefore(text.buildParagraph(doc, spec, model), blocks[0] ?? null)
    }

    for (const el of blocks) {
      xml.remove(el)
    }

    this.package.touch(part)
  }

  private notes(): Note[] {
    const part = this.package.related(this.main, REL.footnotes)[0]
    if (!part) {
      return []
    }

    return xml
      .children(this.package.root(part), W, 'footnote')
      .filter((el) => !NOTE_TYPES.includes(text.wGet(el, 'type')))
      .map((el) => ({ id: text.wGet(el, 'id'), text: textOf(el) }))
  }

  private comments(): Comment[] {
    const part = this.package.related(this.main, REL.comments)[0]
    if (!part) {
      return []
    }

    return xml.children(this.package.root(part), W, 'comment').map((el) => ({
      id: text.wGet(el, 'id'),
      author: text.wGet(el, 'author') || null,
      text: textOf(el)
    }))
  }

  private templateBefore(blocks: Element[], at: number): Element | null {
    for (let i = at - 1; i >= 0; i--) {
      const el = blocks[i]
      if (el.localName !== 'p' || this.headingLevel(el) !== null) {
        return null
      }

      if (text.paragraphText(el) === '' && !xml.descendants(el, W, 'r').length) {
        continue
      }

      return el
    }

    return null
  }

  private blankOnly(blocks: Element[]): Element | null {
    if (blocks.length !== 1 || blocks[0].localName !== 'p') return null
    return isEmptyParagraph(blocks[0]) ? blocks[0] : null
  }

  private ensureTrailingParagraph(): void {
    const blocks = this.blocks()
    const last = blocks[blocks.length - 1]
    if (!last || last.localName !== 'p') {
      const p = text.wEl(this.doc(), 'p')
      this.stamp([p])
      this.body().insertBefore(p, this.sectPr())
    }
  }

  private styleIdOf(p: Element): string | null {
    const pPr = xml.child(p, W, 'pPr')
    const pStyle = pPr ? xml.child(pPr, W, 'pStyle') : null
    return pStyle ? pStyle.getAttributeNS(W, 'val') : null
  }

  private headingLevel(el: Element): number | null {
    if (el.localName !== 'p') return null
    const pPr = xml.child(el, W, 'pPr')
    const outline = pPr ? xml.child(pPr, W, 'outlineLvl') : null
    if (outline) {
      const level = Number(text.wGet(outline, 'val'))
      return level < 9 ? level + 1 : null
    }

    return this.styles.headingLevel(this.styleIdOf(el))
  }

  private isList(p: Element): boolean {
    const pPr = xml.child(p, W, 'pPr')
    const numPr = pPr ? xml.child(pPr, W, 'numPr') : null
    const numId = numPr ? text.wGet(xml.child(numPr, W, 'numId'), 'val') : ''
    if (numId) {
      return numId !== '0'
    }

    return this.styles.isList(this.styleIdOf(p))
  }

  private heading(blocks: Element[], index: number): number {
    this.block(index)
    const level = this.headingLevel(blocks[index])
    if (level === null || level < 1) {
      const headings = this.sections().map((s) => `${s.index} "${s.title}"`)
      throw new DocsError(
        'NOT_HEADING',
        `block ${index} is not a heading; the headings are ${headings.join(', ') || 'none'}`
      )
    }

    return level
  }

  private sectionEnd(blocks: Element[], start: HeadingAt): number {
    for (let i = start.index + 1; i < blocks.length; i++) {
      const other = this.headingLevel(blocks[i])
      if (other !== null && other <= start.level) {
        return i
      }
    }

    return blocks.length
  }

  private describe(el: Element, index: number): Block {
    if (el.localName === 'tbl') {
      return { index, ...ids.idField(el), type: 'table', rows: tableRows(el) }
    }

    if (el.localName === 'sdt') {
      return { index, ...ids.idField(el), type: 'other', text: text.blockText(el) }
    }

    const out: ParagraphBlock = {
      index,
      ...ids.idField(el),
      type: 'paragraph',
      style: this.styles.nameOf(this.styleIdOf(el)),
      text: text.paragraphText(el)
    }
    const level = this.headingLevel(el)
    if (level !== null) {
      out.heading = level
    }

    if (this.isList(el)) {
      out.list = true
      out.level = listLevel(el)
    }

    describeExtras(el, out)
    const charts = this.chartParts(el)
    if (charts.length) {
      out.charts = charts.map((part) => chart.readChart(this.package.xml(part)))
    }

    if (text.hasPageBreak(el)) {
      out.pageBreak = true
    }

    return out
  }

  private readBlock(el: Element, index: number, runs: boolean): Block {
    const block = this.describe(el, index)
    if (runs && block.type === 'paragraph') {
      block.runs = readRuns(el, (link) => this.linkTarget(link))
    }

    return block
  }

  private chartParts(el: Element): string[] {
    const rels = this.package.rels(this.main)
    const out: string[] = []
    for (const ref of xml.descendants(el, chart.C, 'chart')) {
      const rel = rels.get(ref.getAttributeNS(R, 'id'))
      if (rel && !rel.external && rel.part && this.package.has(rel.part)) {
        out.push(rel.part)
      }
    }

    return out
  }

  private chartPart(index: number, which: number): string {
    const parts = this.chartParts(this.block(index))
    if (parts.length === 0) {
      throw new DocsError('NOT_CHART', `block ${index} holds no chart`)
    }

    const part = parts[which]
    if (!part) {
      throw new DocsError(
        'NOT_CHART',
        `block ${index} holds ${parts.length} chart(s); chart ${which} does not exist`
      )
    }

    return part
  }

  private partTexts(type: string): string[] {
    const out: string[] = []
    for (const part of this.package.related(this.main, type)) {
      const value = textOf(this.package.root(part))
      if (value && !out.includes(value)) {
        out.push(value)
      }
    }

    return out
  }

  private textWidth(): number {
    const sectPr = this.sectPr()
    const pgSz = sectPr ? xml.child(sectPr, W, 'pgSz') : null
    const pgMar = sectPr ? xml.child(sectPr, W, 'pgMar') : null
    const width = pgSz ? Number(text.wGet(pgSz, 'w')) : 12240
    const left = pgMar ? Number(text.wGet(pgMar, 'left')) : 1440
    const right = pgMar ? Number(text.wGet(pgMar, 'right')) : 1440
    const usable = width - left - right
    return usable > 0 ? usable : 9360
  }

  private imageSize(info: ImageInfo, opts: ImageOptions): Size {
    checkPositive({ width: opts.width, height: opts.height })
    const ratio = info.height / info.width
    if (opts.width !== undefined && opts.height !== undefined) {
      return { width: opts.width, height: opts.height }
    }

    if (opts.width !== undefined) {
      return { width: opts.width, height: Math.round(opts.width * ratio) }
    }

    if (opts.height !== undefined) {
      return { width: Math.round(opts.height / ratio), height: opts.height }
    }

    const width = Math.min(info.width * EMU_PER_PIXEL, this.textWidth() * EMU_PER_TWIP)

    return { width, height: Math.round(width * ratio) }
  }

  private nextDocPrId(): number {
    const parts = [
      this.main,
      ...DRAWING_PARTS.flatMap((type) => this.package.related(this.main, type))
    ]
    let max = 0
    for (const part of parts) {
      for (const el of xml.descendants(this.package.xml(part), WP, 'docPr')) {
        max = Math.max(max, Number(el.getAttribute('id')) || 0)
      }
    }

    return max + 1
  }

  private tableStyleId(style: string | undefined): string | null {
    if (style !== undefined) {
      return this.styles.require(style, 'table').id
    }

    return this.styles.resolve('Table Grid', 'table')?.id ?? null
  }

  private tableElement(spec: TableSpec): Element {
    const doc = this.doc()
    const cellWidth = Math.floor(this.textWidth() / spec.columns)
    const tbl = text.wEl(doc, 'tbl')
    tbl.appendChild(tableProperties(doc, spec))
    const grid = xml.append(tbl, text.wEl(doc, 'tblGrid'))
    for (let c = 0; c < spec.columns; c++) {
      grid.appendChild(text.wEl(doc, 'gridCol', { w: cellWidth }))
    }

    spec.rows.forEach((row, r) =>
      tbl.appendChild(tableRow(doc, { row, first: r === 0, spec, cellWidth }))
    )

    return tbl
  }

  private indexOf(ref: BlockRef): number {
    return ids.resolve(this.blocks(), ref)
  }

  private at(where: BlockRef, blocks: Element[]): number {
    const at = ids.resolve(blocks, where)
    this.position(at, blocks.length)
    return at
  }

  private stamp(blocks: Element[]): string[] {
    return ids.stamp(this.package.root(this.main), blocks)
  }

  private targets(ref: BlockRef | undefined): Target[] {
    if (ref === undefined) {
      return this.blocks().map((el, index) => [index, el])
    }

    const index = this.indexOf(ref)

    return [[index, this.block(index)]]
  }

  private touchIf(changed: number): void {
    if (changed > 0) {
      this.touch()
    }
  }

  private linkTarget(link: Element): string | undefined {
    const anchor = text.wGet(link, 'anchor')
    if (anchor) {
      return `#${anchor}`
    }

    return this.package.rels(this.main).get(link.getAttributeNS(R, 'id'))?.target
  }

  private summary(el: Element): finder.Summary {
    const paragraph = el.localName === 'p'
    const styleId = paragraph ? this.styleIdOf(el) : finder.tableStyleId(el)
    return {
      type: finder.blockType(el),
      style: paragraph || styleId ? this.styles.nameOf(styleId) : '',
      styleId,
      heading: this.headingLevel(el),
      text: text.blockText(el)
    }
  }

  private insertElements(at: number, built: Element[]): InsertResult {
    const blocks = this.blocks()
    const { anchor, index, trailer } = this.anchor(blocks, at)
    const blank = this.blankOnly(blocks)
    const ids = this.stamp(built)
    for (const el of built) {
      this.body().insertBefore(el, anchor)
    }

    if (blank) {
      xml.remove(blank)
    }

    if (trailer) {
      xml.remove(trailer)
    }

    this.ensureTrailingParagraph()
    this.touch()

    return { index: index - (blank && index > 0 ? 1 : 0), count: built.length, ids }
  }

  private resolveChanges(opts: ChangeOptions, mode: 'accept' | 'reject'): ChangeCounts {
    const scope = opts.in === undefined ? this.body() : this.block(this.indexOf(opts.in))
    const counts = revisions.resolve([scope], mode)
    if (revisions.total(counts)) {
      this.touch()
    }

    return counts
  }
}

interface TableSpec {
  rows: (string | number | null)[][]
  columns: number
  header: boolean
  styleId: string | null
}

interface RowBuild {
  row: (string | number | null)[]
  first: boolean
  spec: TableSpec
  cellWidth: number
}

function tableProperties(doc: XmlDocument, spec: TableSpec): Element {
  const tblPr = text.wEl(doc, 'tblPr')
  if (spec.styleId) {
    tblPr.appendChild(text.wEl(doc, 'tblStyle', { val: spec.styleId }))
  }

  tblPr.appendChild(text.wEl(doc, 'tblW', { w: 0, type: 'auto' }))
  if (!spec.styleId) {
    tblPr.appendChild(borders(doc))
  }

  tblPr.appendChild(
    text.wEl(doc, 'tblLook', {
      val: '04A0',
      firstRow: spec.header ? 1 : 0,
      lastRow: 0,
      firstColumn: 1,
      lastColumn: 0,
      noHBand: 0,
      noVBand: 1
    })
  )

  return tblPr
}

function tableRow(doc: XmlDocument, build: RowBuild): Element {
  const headerRow = build.spec.header && build.first
  const tr = text.wEl(doc, 'tr')
  if (headerRow) {
    xml.append(tr, text.wEl(doc, 'trPr')).appendChild(text.wEl(doc, 'tblHeader'))
  }

  for (let c = 0; c < build.spec.columns; c++) {
    const tc = xml.append(tr, text.wEl(doc, 'tc'))
    xml
      .append(tc, text.wEl(doc, 'tcPr'))
      .appendChild(text.wEl(doc, 'tcW', { w: build.cellWidth, type: 'dxa' }))
    const value = cellText(build.row[c])
    tc.appendChild(
      text.buildParagraph(doc, headerRow ? { runs: [{ text: value, bold: true }] } : value, null)
    )
  }

  return tr
}

function fillCell(tc: Element, value: string): void {
  const tcPr = xml.child(tc, W, 'tcPr')
  for (const v of tcPr ? xml.children(tcPr, W, 'vMerge') : []) {
    xml.remove(v)
  }

  text.replaceContent(clearCell(tc), value)
}

function checkMerge(tc: Element, ref: CellRef): Element {
  const tcPr = xml.child(tc, W, 'tcPr')
  const vMerge = tcPr ? xml.child(tcPr, W, 'vMerge') : null
  if (vMerge && text.wGet(vMerge, 'val') !== 'restart') {
    throw new DocsError(
      'CELL_MERGED',
      `cell ${ref.row},${ref.column} is covered by a merged cell above it; write that cell instead`
    )
  }

  return tc
}

function headerParagraphs(value: unknown, kind: string): unknown[] {
  if (typeof value === 'string') {
    return value.split('\n')
  }

  if (Array.isArray(value) && value.length) return value

  throw new DocsError('INVALID_TEXT', `${kind} text must be a string or an array of paragraphs`)
}

function headerType(opts: HeaderOptions): 'default' | 'first' {
  const type = opts.type ?? 'default'
  if (type !== 'default' && type !== 'first') {
    throw new DocsError('INVALID_OPTIONS', 'type must be "default" or "first"')
  }

  return type
}

function tableRows(el: Element): string[][] {
  return unwrapped(el, 'tr').map((tr) =>
    unwrapped(tr, 'tc').map((tc) =>
      xml
        .children(tc, W)
        .filter((n) => CELL_BLOCKS.has(n.localName ?? ''))
        .map(text.blockText)
        .join('\n')
    )
  )
}

function detailRuns(detail: unknown): boolean {
  if (detail === undefined) return false
  if (detail === 'runs') return true
  throw new DocsError('INVALID_OPTIONS', `detail must be "runs", got ${String(detail)}`, {
    field: 'detail',
    expected: ['runs'],
    got: detail
  })
}

function listLevel(el: Element): number {
  const pPr = xml.child(el, W, 'pPr')
  const numPr = pPr ? xml.child(pPr, W, 'numPr') : null
  const ilvl = numPr ? xml.child(numPr, W, 'ilvl') : null
  return ilvl ? Number(text.wGet(ilvl, 'val')) : 0
}

function describeExtras(el: Element, out: ParagraphBlock): void {
  const footnotes = xml.descendants(el, W, 'footnoteReference').map((r) => text.wGet(r, 'id'))
  if (footnotes.length) {
    out.footnotes = footnotes
  }

  const anchors = [
    ...xml.descendants(el, W, 'commentRangeStart'),
    ...xml.descendants(el, W, 'commentReference')
  ]
  const comments = [...new Set(anchors.map((r) => text.wGet(r, 'id')))]
  if (comments.length) {
    out.comments = comments
  }

  const pictures = xml.descendants(el, PIC, 'pic').length
  if (pictures) {
    out.images = pictures
  }
}
