import * as xml from '../xml.ts'
import * as text from './text.ts'
import { DocsError } from '../errors.ts'
import { ListAllocator, place } from './lists.ts'
import { parse, type Inline, type MarkdownBlock } from './markdown.ts'
import { W, R, REL_BASE, addExternal, styleId, textWidth, runStyle, type Host } from './host.ts'
import type { Document, Element } from '../xml.ts'
import type { Run } from './text.ts'

type TableBlock = Extract<MarkdownBlock, { type: 'table' }>
type ItemBlock = Extract<MarkdownBlock, { type: 'item' }>

interface RowSpec {
  columns: number
  header: boolean
}

interface Builder {
  host: Host
  doc: Document
  lists: ListAllocator
  links: Map<string, string>
}

export const CODE_FONT = 'Consolas'
const LINK_COLOR = '0563C1'

function runOf(inline: Inline, link: boolean): Run {
  const run: Run = { text: inline.text }
  if (inline.bold) {
    run.bold = true
  }

  if (inline.italic) {
    run.italic = true
  }

  if (inline.code) {
    run.font = CODE_FONT
  }

  if (link) {
    run.color = LINK_COLOR
    run.underline = true
  }

  return run
}

function linkId(b: Builder, target: string): string {
  const known = b.links.get(target)
  if (known) {
    return known
  }

  const id = addExternal(b.host, REL_BASE + 'hyperlink', target)
  b.links.set(target, id)

  return id
}

interface LinkRun {
  r: Element
  target: string
  previous: Element | null
}

function wrapLink(b: Builder, { r, target, previous }: LinkRun): Element {
  if (previous && previous.getAttributeNS(R, 'id') === linkId(b, target)) {
    previous.appendChild(r)

    return previous
  }

  const link = text.wEl(b.doc, 'hyperlink', { history: 1 })
  xml.setNsAttribute(link, { namespace: R, local: 'id' }, linkId(b, target))
  r.parentNode?.insertBefore(link, r)
  link.appendChild(r)

  return link
}

function styleLinkRun(b: Builder, r: Element): void {
  const rPr = xml.child(r, W, 'rPr')
  if (!rPr || !runStyle(b.host, r, 'Hyperlink')) {
    return
  }

  for (const local of ['color', 'u']) {
    for (const el of xml.children(rPr, W, local)) {
      xml.remove(el)
    }
  }
}

function fill(b: Builder, p: Element, parts: Inline[]): Element {
  const runs = parts.map((inline) => runOf(inline, inline.link !== undefined))
  const built = text.buildParagraph(b.doc, { runs: runs.length ? runs : [''] }, null)
  let link: Element | null = null
  xml.children(built, W, 'r').forEach((r, i) => {
    const target = parts[i]?.link
    p.appendChild(r)
    if (target === undefined) {
      link = null
      return
    }

    styleLinkRun(b, r)
    link = wrapLink(b, { r, target, previous: link })
  })

  return p
}

function paragraph(b: Builder, style: string | null): Element {
  const p = text.wEl(b.doc, 'p')
  if (style) {
    text.ensure(p, 'pPr').appendChild(text.wEl(b.doc, 'pStyle', { val: style }))
  }

  return p
}

function heading(b: Builder, level: number, parts: Inline[]): Element {
  const style = styleId(b.host, `heading ${level}`)
  if (style) {
    return fill(b, paragraph(b, style), parts)
  }

  const p = paragraph(b, null)
  xml.insertOrdered(
    text.ensure(p, 'pPr'),
    text.wEl(b.doc, 'outlineLvl', { val: level - 1 }),
    text.PARAGRAPH_ORDER
  )

  return fill(
    b,
    p,
    parts.map((inline) => ({ ...inline, bold: true }))
  )
}

function rule(b: Builder): Element {
  const p = paragraph(b, null)
  const border = xml.insertOrdered(
    text.ensure(p, 'pPr'),
    text.wEl(b.doc, 'pBdr'),
    text.PARAGRAPH_ORDER
  )
  border.appendChild(text.wEl(b.doc, 'bottom', { val: 'single', sz: 6, space: 1, color: 'auto' }))

  return p
}

function item(b: Builder, block: ItemBlock): Element {
  const p = fill(b, paragraph(b, styleId(b.host, 'List Paragraph')), block.inlines)
  place(p, { numId: b.lists.numId(block.ordered ? 'number' : 'bullet'), level: block.level })
  return p
}

function tableProperties(b: Builder): Element {
  const tblPr = text.wEl(b.doc, 'tblPr')
  const style = styleId(b.host, 'Table Grid', 'table')
  if (style) {
    tblPr.appendChild(text.wEl(b.doc, 'tblStyle', { val: style }))
  }

  tblPr.appendChild(text.wEl(b.doc, 'tblW', { w: 0, type: 'auto' }))
  if (!style) {
    const borders = xml.append(tblPr, text.wEl(b.doc, 'tblBorders'))
    for (const side of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) {
      borders.appendChild(text.wEl(b.doc, side, { val: 'single', sz: 4, space: 0, color: 'auto' }))
    }
  }

  tblPr.appendChild(
    text.wEl(b.doc, 'tblLook', {
      val: '04A0',
      firstRow: 1,
      lastRow: 0,
      firstColumn: 1,
      lastColumn: 0,
      noHBand: 0,
      noVBand: 1
    })
  )

  return tblPr
}

function row(b: Builder, cells: Inline[][], spec: RowSpec): Element {
  const tr = text.wEl(b.doc, 'tr')
  if (spec.header) {
    xml.append(tr, text.wEl(b.doc, 'trPr')).appendChild(text.wEl(b.doc, 'tblHeader'))
  }

  const width = Math.floor(textWidth(b.host) / spec.columns)
  for (let c = 0; c < spec.columns; c++) {
    const tc = xml.append(tr, text.wEl(b.doc, 'tc'))
    xml
      .append(tc, text.wEl(b.doc, 'tcPr'))
      .appendChild(text.wEl(b.doc, 'tcW', { w: width, type: 'dxa' }))
    const parts = cells[c] ?? []
    tc.appendChild(
      fill(b, paragraph(b, null), spec.header ? parts.map((p) => ({ ...p, bold: true })) : parts)
    )
  }

  return tr
}

function table(b: Builder, block: TableBlock): Element {
  const columns = Math.max(block.header.length, ...block.rows.map((r) => r.length))
  const tbl = text.wEl(b.doc, 'tbl')
  tbl.appendChild(tableProperties(b))
  const grid = xml.append(tbl, text.wEl(b.doc, 'tblGrid'))
  for (let c = 0; c < columns; c++) {
    grid.appendChild(text.wEl(b.doc, 'gridCol', { w: Math.floor(textWidth(b.host) / columns) }))
  }

  tbl.appendChild(row(b, block.header, { columns, header: true }))
  for (const cells of block.rows) {
    tbl.appendChild(row(b, cells, { columns, header: false }))
  }

  return tbl
}

function code(b: Builder, lines: string[]): Element[] {
  return (lines.length ? lines : ['']).map((line) =>
    fill(b, paragraph(b, null), [{ text: line, code: true }])
  )
}

function element(b: Builder, block: MarkdownBlock): Element[] {
  if (block.type !== 'item') {
    b.lists.end()
  }

  switch (block.type) {
    case 'heading':
      return [heading(b, block.level, block.inlines)]
    case 'paragraph':
      return [fill(b, paragraph(b, null), block.inlines)]
    case 'item':
      return [item(b, block)]
    case 'table':
      return [table(b, block)]
    case 'rule':
      return [rule(b)]
    case 'code':
      return code(b, block.lines)
  }
}

export function build(host: Host, markdown: unknown): Element[] {
  if (typeof markdown !== 'string') {
    throw new DocsError('INVALID_TEXT', 'markdown must be a string', {
      field: 'markdown',
      expected: 'string',
      got: typeof markdown
    })
  }

  const blocks = parse(markdown)
  if (blocks.length === 0) {
    throw new DocsError('INVALID_TEXT', 'the markdown holds no blocks to insert')
  }

  const doc = host.package.xml(host.main)
  const b: Builder = { host, doc, lists: new ListAllocator(host), links: new Map() }

  return blocks.flatMap((block) => element(b, block))
}
