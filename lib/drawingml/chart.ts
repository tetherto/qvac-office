import * as xml from '../xml.ts'
import * as paths from '../opc/paths.ts'
import * as text from './text.ts'
import * as refs from '../xlsx/refs.ts'
import { writeWorkbook } from '../xlsx/write.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Document, Element, Build } from '../xml.ts'
import type { Package } from '../opc/package.ts'

export type ChartType = 'column' | 'bar' | 'line' | 'pie' | 'doughnut' | 'area'
export type Legend = false | 'right' | 'left' | 'top' | 'bottom'
export type Stacking = boolean | 'percent'
export type Category = string | number

export interface Series {
  name: string
  values: (number | null)[]
}

export interface ChartData {
  categories: Category[]
  series: Series[]
}

export interface ChartSpec extends ChartData {
  type: ChartType
  title?: string
  stacked?: Stacking
  legend?: Legend
}

export interface ReadSeries {
  name: string | null
  values: (string | number | null)[]
}

export interface ChartInfo {
  type: string | null
  types?: string[]
  title: string | null
  categories: (string | number | null)[]
  series: ReadSeries[]
}

export interface SeriesLink {
  name: string | null
  values: string
}

export interface ChartLinks {
  categories: string
  series: SeriesLink[]
}

export interface WriteOptions {
  links?: ChartLinks
}

export interface ChartRef {
  pkg: Package
  part: string
}

export interface DataResult {
  series: number
  categories: number
}

interface SeriesWrite {
  index: number
  data: ChartData
  links?: ChartLinks
}

interface Cached {
  values: (string | number | null)[]
  formatCode?: string
}

interface Reference {
  formula: string
  cache: Element
}

export const C = xml.ns.c
export const CHART_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
export const GRAPHIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
export const CREATABLE: readonly ChartType[] = ['column', 'bar', 'line', 'pie', 'doughnut', 'area']

const A = xml.ns.a
const R = xml.ns.r
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const REL_PACKAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package'
const SHEET = 'Sheet1'
const LEGEND_POSITIONS = { right: 'r', left: 'l', top: 't', bottom: 'b' } as const
const EDITABLE = new Set([
  'barChart',
  'bar3DChart',
  'lineChart',
  'line3DChart',
  'pieChart',
  'pie3DChart',
  'doughnutChart',
  'areaChart',
  'area3DChart',
  'radarChart'
])
const SERIES_ORDER = [
  'idx',
  'order',
  'tx',
  'spPr',
  'invertIfNegative',
  'pictureOptions',
  'marker',
  'explosion',
  'dPt',
  'dLbls',
  'trendline',
  'errBars',
  'cat',
  'val',
  'smooth',
  'shape',
  'bubble3D',
  'extLst'
]
const CHART_ORDER = [
  'title',
  'autoTitleDeleted',
  'pivotFmts',
  'view3D',
  'floor',
  'sideWall',
  'backWall',
  'plotArea',
  'legend',
  'plotVisOnly',
  'dispBlanksAs',
  'showDLblsOverMax',
  'extLst'
]
const SPACE_ORDER = [
  'date1904',
  'lang',
  'roundedCorners',
  'AlternateContent',
  'style',
  'clrMapOvr',
  'pivotSource',
  'protection',
  'chart',
  'spPr',
  'txPr',
  'externalData',
  'printSettings',
  'userShapes',
  'extLst'
]

function groupsOf(doc: Document): Element[] {
  const plotArea = xml.descendants(doc, C, 'plotArea')[0]
  if (!plotArea) {
    return []
  }

  return xml.children(plotArea, C).filter((el) => (el.localName ?? '').endsWith('Chart'))
}

function typeOf(group: Element): string {
  const name = group.localName ?? ''
  if (name === 'barChart' || name === 'bar3DChart') {
    return xml.child(group, C, 'barDir')?.getAttribute('val') === 'bar' ? 'bar' : 'column'
  }

  if (name === 'ofPieChart') {
    return 'pie-of-pie'
  }

  return name.replace(/3D/, '').replace(/Chart$/, '')
}

function seriesOf(doc: Document): Element[] {
  return groupsOf(doc).flatMap((group) => xml.children(group, C, 'ser'))
}

function cacheOf(container: Element): Element | undefined {
  for (const local of ['strCache', 'numCache', 'strLit', 'numLit', 'lvl']) {
    const found = xml.descendants(container, C, local)[0]
    if (found) {
      return found
    }
  }

  return undefined
}

function points(container: Element | null): (string | number | null)[] {
  if (!container) return []
  const cache = cacheOf(container)
  if (!cache) {
    const v = xml.child(container, C, 'v')
    return v ? [xml.text(v)] : []
  }

  const numeric = cache.localName === 'numCache' || cache.localName === 'numLit'
  const out: (string | number | null)[] = new Array(
    Number(xml.child(cache, C, 'ptCount')?.getAttribute('val')) || 0
  ).fill(null)
  for (const pt of xml.children(cache, C, 'pt')) {
    const idx = Number(pt.getAttribute('idx'))
    const v = xml.child(pt, C, 'v')
    if (!v || !(idx >= 0)) {
      continue
    }

    while (out.length <= idx) {
      out.push(null)
    }

    out[idx] = numeric ? Number(xml.text(v)) : xml.text(v)
  }

  return out
}

function readTitle(doc: Document): string | null {
  const chart = xml.child(xml.root(doc), C, 'chart')
  const title = chart ? xml.child(chart, C, 'title') : null
  const tx = title ? xml.child(title, C, 'tx') : null
  if (!tx) {
    return null
  }

  const rich = xml.child(tx, C, 'rich')
  if (rich) {
    return text.readText(rich)
  }

  const name = points(tx)

  return name.length ? String(name[0]) : null
}

function nameOf(ser: Element): string | null {
  const first = points(xml.child(ser, C, 'tx'))[0]
  return first === undefined || first === null ? null : String(first)
}

export function readChart(doc: Document): ChartInfo {
  const groups = groupsOf(doc)
  const sers = seriesOf(doc)
  const first = sers[0]
  let type: string | null = null
  if (groups.length === 1) {
    type = typeOf(groups[0])
  } else if (groups.length > 1) {
    type = 'combo'
  }

  const out: ChartInfo = {
    type,
    title: readTitle(doc),
    categories: first ? points(xml.child(first, C, 'cat') ?? xml.child(first, C, 'xVal')) : [],
    series: sers.map((ser) => ({
      name: nameOf(ser),
      values: points(xml.child(ser, C, 'val') ?? xml.child(ser, C, 'yVal'))
    }))
  }
  if (groups.length > 1) {
    out.types = groups.map(typeOf)
  }

  return out
}

function isRecordLike(value: unknown): value is Record<string, unknown> {
  return isRecord(value) && !Array.isArray(value)
}

function validateCategories(categories: unknown): Category[] {
  if (!Array.isArray(categories) || categories.length === 0) {
    throw new DocsError('INVALID_CHART_DATA', 'categories must be a non-empty array')
  }

  if (categories.length > 10000) {
    throw new DocsError('INVALID_CHART_DATA', 'a chart holds at most 10000 categories')
  }

  return categories.map((c, i) => {
    if (typeof c === 'string' || (typeof c === 'number' && Number.isFinite(c))) {
      return c
    }

    throw new DocsError('INVALID_CHART_DATA', `category ${i} must be a string or a number`)
  })
}

function validateSeries(s: unknown, i: number, count: number): Series {
  if (!isRecordLike(s) || typeof s.name !== 'string' || !Array.isArray(s.values)) {
    throw new DocsError('INVALID_CHART_DATA', `series ${i} must be { name, values }`)
  }

  const name = s.name
  if (s.values.length !== count) {
    throw new DocsError(
      'INVALID_CHART_DATA',
      `series ${i} ("${name}") has ${s.values.length} values but there are ${count} categories`
    )
  }

  const values = s.values.map((v, j) => {
    if (v === null || (typeof v === 'number' && Number.isFinite(v))) return v
    throw new DocsError('INVALID_CHART_DATA', `value ${j} of series ${i} must be a number or null`)
  })

  return { name, values }
}

export function validateData(data: unknown): ChartData {
  if (!isRecordLike(data)) {
    throw new DocsError('INVALID_CHART_DATA', 'chart data must be { categories, series }')
  }

  const categories = validateCategories(data.categories)
  const { series } = data
  if (!Array.isArray(series) || series.length === 0) {
    throw new DocsError(
      'INVALID_CHART_DATA',
      'series must be a non-empty array of { name, values }'
    )
  }

  if (series.length > 255) {
    throw new DocsError('INVALID_CHART_DATA', 'a chart holds at most 255 series')
  }

  return { categories, series: series.map((s, i) => validateSeries(s, i, categories.length)) }
}

function valEl(c: Build, local: string, value: string | number): Element {
  return c(local, { val: value })
}

function textEl(c: Build, local: string, value: string): Element {
  const el = c(local)
  xml.setText(el, value)
  return el
}

function cacheEl(c: Build, local: string, cached: Cached): Element {
  const cache = c(local)
  if (cached.formatCode !== undefined) {
    cache.appendChild(textEl(c, 'formatCode', cached.formatCode))
  }

  cache.appendChild(valEl(c, 'ptCount', cached.values.length))
  cached.values.forEach((value, idx) => {
    if (value === null || value === undefined) return
    xml.append(cache, c('pt', { idx })).appendChild(textEl(c, 'v', String(value)))
  })

  return cache
}

function refEl(c: Build, local: string, ref: Reference): Element {
  const el = c(local)
  el.appendChild(textEl(c, 'f', ref.formula))
  el.appendChild(ref.cache)
  return el
}

function replaceChild(
  ser: Element,
  local: string,
  fill: (container: Element, formatCode: string) => void
): void {
  const c = xml.builder(xml.owner(ser), C)
  const container = xml.child(ser, C, local) ?? xml.insertOrdered(ser, c(local), SERIES_ORDER)
  const previous = xml.descendants(container, C, 'formatCode')[0]
  while (container.firstChild) {
    container.removeChild(container.firstChild)
  }

  fill(container, previous ? xml.text(previous) : 'General')
}

interface SeriesFormulas extends SeriesLink {
  categories: string
}

function formulas(write: SeriesWrite): SeriesFormulas {
  const count = write.data.categories.length
  const link = write.links?.series[write.index]
  if (write.links && link) {
    return { name: link.name, values: link.values, categories: write.links.categories }
  }

  const col = write.index + 1

  return {
    name: refs.absolute(SHEET, { top: 0, left: col, bottom: 0, right: col }),
    values: refs.absolute(SHEET, { top: 1, left: col, bottom: count, right: col }),
    categories: refs.absolute(SHEET, { top: 1, left: 0, bottom: count, right: 0 })
  }
}

function writeSeries(ser: Element, write: SeriesWrite): void {
  const c = xml.builder(xml.owner(ser), C)
  const { data, index } = write
  const series = data.series[index]
  const f = formulas(write)
  replaceChild(ser, 'tx', (tx) => {
    const name = f.name
    if (name) {
      tx.appendChild(
        refEl(c, 'strRef', {
          formula: name,
          cache: cacheEl(c, 'strCache', { values: [series.name] })
        })
      )
    } else {
      tx.appendChild(textEl(c, 'v', series.name))
    }
  })
  const numeric = data.categories.every((cat) => typeof cat === 'number')
  replaceChild(ser, 'cat', (cat, formatCode) => {
    const cache = numeric
      ? cacheEl(c, 'numCache', { values: data.categories, formatCode })
      : cacheEl(c, 'strCache', { values: data.categories.map(String) })
    cat.appendChild(refEl(c, numeric ? 'numRef' : 'strRef', { formula: f.categories, cache }))
  })
  replaceChild(ser, 'val', (val, formatCode) => {
    val.appendChild(
      refEl(c, 'numRef', {
        formula: f.values,
        cache: cacheEl(c, 'numCache', { values: series.values, formatCode })
      })
    )
  })
  for (const dPt of xml.children(ser, C, 'dPt')) {
    if (Number(xml.child(dPt, C, 'idx')?.getAttribute('val')) >= data.categories.length) {
      xml.remove(dPt)
    }
  }
}

function setIndex(ser: Element, i: number): void {
  for (const local of ['idx', 'order']) {
    xml.child(ser, C, local)?.setAttribute('val', String(i))
  }
}

function newSeries(doc: Document, group: Element): Element {
  const c = xml.builder(doc, C)
  const ser = c('ser')
  ser.appendChild(valEl(c, 'idx', 0))
  ser.appendChild(valEl(c, 'order', 0))
  if (group.localName === 'barChart') {
    ser.appendChild(valEl(c, 'invertIfNegative', 0))
  }

  if (group.localName === 'lineChart') {
    ser.appendChild(valEl(c, 'smooth', 0))
  }

  return ser
}

function insertSeries(group: Element, ser: Element): void {
  const existing = xml.children(group, C, 'ser')
  const after = existing.length
    ? existing[existing.length - 1]
    : (xml.child(group, C, 'varyColors') ?? xml.child(group, C, 'grouping'))
  group.insertBefore(ser, after ? after.nextSibling : group.firstChild)
}

function embeddedRows(data: ChartData): (string | number | null)[][] {
  return [
    ['', ...data.series.map((s) => s.name)],
    ...data.categories.map((cat, i) => [cat, ...data.series.map((s) => s.values[i])])
  ]
}

function writeEmbedding(ref: ChartRef, data: ChartData): void {
  const { pkg, part } = ref
  const top = pkg.root(part)
  const bytes = writeWorkbook(embeddedRows(data), SHEET)
  const rels = pkg.rels(part)
  let external = xml.child(top, C, 'externalData')
  const current = external ? rels.get(external.getAttributeNS(R, 'id')) : null
  if (
    current &&
    !current.external &&
    current.part &&
    pkg.has(current.part) &&
    current.part.endsWith('.xlsx')
  ) {
    pkg.write(current.part, bytes)
    return
  }

  if (current) {
    rels.remove(current.id)
  }

  pkg.contentTypes.addDefault('xlsx', XLSX_TYPE)
  const embedding = pkg.freeName(
    paths.resolve(part, '../embeddings/Microsoft_Excel_Worksheet%d.xlsx')
  )
  pkg.add(embedding, bytes)
  const rId = rels.add(REL_PACKAGE, embedding)
  if (!external) {
    const c = xml.builder(xml.owner(top), C)
    external = xml.insertOrdered(top, c('externalData'), SPACE_ORDER)
    external.appendChild(valEl(c, 'autoUpdate', 0))
  }

  xml.setNsAttribute(external, { namespace: R, local: 'id' }, rId)
}

function writeCombo(sers: Element[], data: ChartData, opts: WriteOptions): void {
  if (data.series.length !== sers.length) {
    throw new DocsError(
      'CHART_COMBO',
      `this combo chart has ${sers.length} series; give exactly ${sers.length} to keep each on its plot`
    )
  }

  sers.forEach((ser, index) => writeSeries(ser, { index, data, links: opts.links }))
}

function writeSingle(group: Element, data: ChartData, opts: WriteOptions): void {
  const doc = xml.owner(group)
  const sers = xml.children(group, C, 'ser')
  for (const extra of sers.splice(data.series.length)) {
    xml.remove(extra)
  }

  while (sers.length < data.series.length) {
    const template = sers[sers.length - 1]
    const ser = template ? xml.clone(template) : newSeries(doc, group)
    for (const local of ['spPr', 'dPt', 'trendline', 'errBars']) {
      for (const el of xml.children(ser, C, local)) {
        xml.remove(el)
      }
    }

    insertSeries(group, ser)
    sers.push(ser)
  }

  sers.forEach((ser, index) => {
    setIndex(ser, index)
    writeSeries(ser, { index, data, links: opts.links })
  })
}

export function setChartData(ref: ChartRef, input: unknown, opts: WriteOptions = {}): DataResult {
  const data = validateData(input)
  const doc = ref.pkg.xml(ref.part)
  const groups = groupsOf(doc)
  if (groups.length === 0) {
    throw new DocsError('UNSUPPORTED_CHART', 'the chart has no plot to write data into')
  }

  const unsupported = groups.find((g) => !EDITABLE.has(g.localName ?? ''))
  if (unsupported) {
    throw new DocsError(
      'UNSUPPORTED_CHART',
      `${typeOf(unsupported)} charts cannot be edited; column, bar, line, pie, doughnut, area and radar charts can`
    )
  }

  if (groups.length > 1) {
    writeCombo(seriesOf(doc), data, opts)
  } else {
    writeSingle(groups[0], data, opts)
  }

  if (!opts.links) {
    writeEmbedding(ref, data)
  }

  ref.pkg.touch(ref.part)

  return { series: data.series.length, categories: data.categories.length }
}

function richTitle(title: Element): Element {
  const doc = xml.owner(title)
  const c = xml.builder(doc, C)
  const a = xml.builder(doc, A)
  const tx = xml.child(title, C, 'tx')
  const existing = tx ? xml.child(tx, C, 'rich') : null
  if (existing) {
    return existing
  }

  if (tx) {
    xml.remove(tx)
  }

  const fresh = c('tx')
  title.insertBefore(fresh, title.firstChild)
  const rich = xml.append(fresh, c('rich'))
  for (const local of ['bodyPr', 'lstStyle', 'p']) {
    rich.appendChild(a(local))
  }

  return rich
}

export interface Changed {
  changed: boolean
}

export function setChartTitle(ref: ChartRef, value: text.TextInput | null): Changed {
  const chart = xml.child(ref.pkg.root(ref.part), C, 'chart')
  if (!chart) {
    throw new DocsError('UNSUPPORTED_CHART', 'the chart part has no chart')
  }

  const c = xml.builder(xml.owner(chart), C)
  const deleted =
    xml.child(chart, C, 'autoTitleDeleted') ??
    xml.insertOrdered(chart, valEl(c, 'autoTitleDeleted', 0), CHART_ORDER)
  let title = xml.child(chart, C, 'title')
  if (value === null || value === '') {
    if (title) {
      xml.remove(title)
    }

    deleted.setAttribute('val', '1')
    ref.pkg.touch(ref.part)

    return { changed: true }
  }

  if (!title) {
    title = xml.insertOrdered(chart, c('title'), CHART_ORDER)
    title.appendChild(valEl(c, 'overlay', 0))
  }

  text.writeParagraphs(richTitle(title), value)
  deleted.setAttribute('val', '0')
  ref.pkg.touch(ref.part)

  return { changed: true }
}

function isChartType(value: unknown): value is ChartType {
  return typeof value === 'string' && CREATABLE.some((type) => type === value)
}

function isLegend(value: unknown): value is Legend {
  return value === false || (typeof value === 'string' && value in LEGEND_POSITIONS)
}

function groupingFor(type: ChartType, stacked: Stacking): string {
  if (stacked === 'percent') return 'percentStacked'
  if (stacked) return 'stacked'
  return type === 'column' || type === 'bar' ? 'clustered' : 'standard'
}

function plotXml(type: ChartType, stacked: Stacking): string {
  const grouping = groupingFor(type, stacked)
  const axes = '<c:axId val="10"/><c:axId val="20"/>'
  if (type === 'column' || type === 'bar') {
    const overlap = stacked ? '<c:overlap val="100"/>' : ''
    return `<c:barChart><c:barDir val="${type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/><c:gapWidth val="150"/>${overlap}${axes}</c:barChart>`
  }

  if (type === 'line') {
    return `<c:lineChart><c:grouping val="${grouping}"/><c:varyColors val="0"/><c:marker val="1"/>${axes}</c:lineChart>`
  }

  if (type === 'area') {
    return `<c:areaChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${axes}</c:areaChart>`
  }

  if (type === 'pie') {
    return '<c:pieChart><c:varyColors val="1"/><c:firstSliceAng val="0"/></c:pieChart>'
  }

  return '<c:doughnutChart><c:varyColors val="1"/><c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>'
}

function axesXml(type: ChartType): string {
  if (!['column', 'bar', 'line', 'area'].includes(type)) return ''
  const catPos = type === 'bar' ? 'l' : 'b'
  const valPos = type === 'bar' ? 'b' : 'l'
  const between = type === 'area' ? 'midCat' : 'between'
  const scaling = '<c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
  const ticks =
    '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'

  return `<c:catAx><c:axId val="10"/>${scaling}<c:axPos val="${catPos}"/>${ticks}<c:crossAx val="20"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx><c:valAx><c:axId val="20"/>${scaling}<c:axPos val="${valPos}"/><c:majorGridlines/>${ticks}<c:crossAx val="10"/><c:crosses val="autoZero"/><c:crossBetween val="${between}"/></c:valAx>`
}

function skeleton(type: ChartType, stacked: Stacking, legend: Legend): string {
  const legendXml =
    legend === false
      ? ''
      : `<c:legend><c:legendPos val="${LEGEND_POSITIONS[legend]}"/><c:overlay val="0"/></c:legend>`
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:date1904 val="0"/><c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>${plotXml(type, stacked)}${axesXml(type)}</c:plotArea>${legendXml}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`
}

export function createChart(ref: ChartRef, spec: unknown, opts: WriteOptions = {}): string {
  if (!isRecordLike(spec) || !isChartType(spec.type)) {
    const type = isRecordLike(spec) ? spec.type : undefined
    throw new DocsError(
      'UNSUPPORTED_CHART',
      `chart type must be one of ${CREATABLE.join(', ')}, got ${JSON.stringify(type)}`
    )
  }

  const type = spec.type
  const data = validateData(spec)
  const stacked = spec.stacked ?? false
  if (stacked !== false && stacked !== true && stacked !== 'percent') {
    throw new DocsError('INVALID_CHART', 'stacked must be false, true or "percent"')
  }

  const legend = spec.legend ?? 'right'
  if (!isLegend(legend)) {
    throw new DocsError('INVALID_CHART', 'legend must be false, "right", "left", "top" or "bottom"')
  }

  const doc = xml.parse(skeleton(type, stacked, legend))
  const group = groupsOf(doc)[0]
  for (let i = 0; i < data.series.length; i++) {
    insertSeries(group, newSeries(doc, group))
  }

  ref.pkg.add(ref.part, doc, CHART_TYPE)
  setChartData(ref, data, opts)
  if (typeof spec.title === 'string' && spec.title) {
    setChartTitle(ref, spec.title)
  }

  return ref.part
}
