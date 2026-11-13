import * as xml from '../xml.ts'
import * as text from './text.ts'
import { DocsError } from '../errors.ts'
import { isBullet } from './lists.ts'
import { W, R, blocks, headingLevel, listInfo, type Host } from './host.ts'
import type { Element, Node } from '../xml.ts'
import type { Inline } from './markdown.ts'

export interface MarkdownOptions {
  from?: number
  count?: number
}

export interface MarkdownResult {
  total: number
  markdown: string
}

interface Line {
  kind: 'item' | 'block'
  value: string
}

const MONOSPACE = new Set([
  'consolas',
  'courier',
  'courier new',
  'menlo',
  'monaco',
  'lucida console',
  'source code pro',
  'cascadia code',
  'cascadia mono'
])
const SKIPPED = new Set(['del', 'moveFrom', 'pPr', 'rPr'])
const CONTAINERS = new Set(['ins', 'moveTo', 'smartTag', 'customXml', 'fldSimple', 'dir', 'bdo'])
const MARKERS = ['', '*', '**', '***']
const INLINE_ESCAPES = /[\\`*_[\]]/g

function on(rPr: Element | null, local: string): boolean {
  const el = rPr ? xml.child(rPr, W, local) : null
  return !!el && !['0', 'false', 'off'].includes(text.wGet(el, 'val'))
}

function runFormat(r: Element): Omit<Inline, 'text'> {
  const rPr = xml.child(r, W, 'rPr')
  const style = text.wGet(rPr ? xml.child(rPr, W, 'rStyle') : null, 'val').toLowerCase()
  const font = text.wGet(rPr ? xml.child(rPr, W, 'rFonts') : null, 'ascii').toLowerCase()
  const out: Omit<Inline, 'text'> = {}
  if (on(rPr, 'b') || style === 'strong') {
    out.bold = true
  }

  if (on(rPr, 'i') || style === 'emphasis') {
    out.italic = true
  }

  if (MONOSPACE.has(font) || /code|verbatim/.test(style)) {
    out.code = true
  }

  return out
}

function linkTarget(host: Host, link: Element): string | undefined {
  const rel = host.package.rels(host.main).get(link.getAttributeNS(R, 'id'))
  if (rel?.external) {
    return rel.target
  }

  const anchor = text.wGet(link, 'anchor')

  return anchor ? `#${anchor}` : undefined
}

interface Walk {
  host: Host
  link: string | undefined
  out: Inline[]
}

function collect(walk: Walk, parent: Element): void {
  for (const el of xml.children(parent, W)) {
    const name = el.localName ?? ''
    if (name === 'r') {
      const value = text.paragraphText(wrapRun(el))
      walk.out.push({ ...runFormat(el), link: walk.link, text: value })
    } else if (name === 'hyperlink') {
      collect({ ...walk, link: linkTarget(walk.host, el) ?? walk.link }, el)
    } else if (name === 'sdt') {
      const content = xml.child(el, W, 'sdtContent')
      if (content) {
        collect(walk, content)
      }
    } else if (CONTAINERS.has(name) && !SKIPPED.has(name)) {
      collect(walk, el)
    }
  }
}

function wrapRun(r: Element): Element {
  const holder = text.wEl(xml.owner(r), 'p')
  holder.appendChild(xml.clone(r))
  return holder
}

function escapeText(value: string): string {
  return value.replace(INLINE_ESCAPES, '\\$&')
}

function codeText(value: string): string {
  const fence = value.includes('`') ? '``' : '`'
  const pad = value.startsWith('`') || value.endsWith('`') ? ' ' : ''
  return `${fence}${pad}${value}${pad}${fence}`
}

function marked(inline: Inline): string {
  if (inline.code) {
    return codeText(inline.text)
  }

  const match = /^(\s*)(.*?)(\s*)$/s.exec(inline.text) ?? ['', '', inline.text, '']
  const core = escapeText(match[2])
  if (!core) {
    return inline.text
  }

  const marker = MARKERS[(inline.bold ? 2 : 0) + (inline.italic ? 1 : 0)]

  return `${match[1]}${marker}${core}${marker}${match[3]}`
}

function encodeTarget(target: string): string {
  return target.replace(
    /[()\s]/g,
    (c) => '%' + c.charCodeAt(0).toString(16).toUpperCase().padStart(2, '0')
  )
}

function joined(inlines: Inline[]): string {
  let out = ''
  for (let i = 0; i < inlines.length;) {
    const link = inlines[i].link
    if (!link) {
      out += marked(inlines[i++])
      continue
    }

    let inner = ''
    for (; i < inlines.length && inlines[i].link === link; i++) {
      inner += marked(inlines[i])
    }

    out += `[${inner}](${encodeTarget(link)})`
  }

  return out
}

function merged(inlines: Inline[]): Inline[] {
  const out: Inline[] = []
  for (const inline of inlines.filter((i) => i.text)) {
    const last = out[out.length - 1]
    const same =
      last &&
      last.bold === inline.bold &&
      last.italic === inline.italic &&
      last.code === inline.code &&
      last.link === inline.link
    if (same) {
      last.text += inline.text
    } else {
      out.push({ ...inline })
    }
  }

  return out
}

export function inlineMarkdown(host: Host, p: Element, plain = false): string {
  const parts: Inline[] = []
  collect({ host, link: undefined, out: parts }, p)
  const flat = parts.map((part) => ({ ...part, text: part.text.replace(/\n/g, ' ') }))
  return joined(merged(plain ? flat.map(({ bold, ...rest }) => rest) : flat))
}

function escapeStart(line: string): string {
  return line
    .replace(/^(#{1,6}\s|[-+*]\s|[>|])/, '\\$1')
    .replace(/^([-_])(?=(\s*\1){2,}\s*$)/, '\\$1')
    .replace(/^(\d+)([.)])/, '$1\\$2')
}

function isRule(p: Element): boolean {
  const pPr = xml.child(p, W, 'pPr')
  const pBdr = pPr ? xml.child(pPr, W, 'pBdr') : null
  return !!pBdr && !!xml.child(pBdr, W, 'bottom') && text.paragraphText(p) === ''
}

function paragraphLine(host: Host, p: Element): Line | null {
  const level = headingLevel(host, p)
  const value = inlineMarkdown(host, p, level !== null)
  if (level !== null && value.trim()) {
    return {
      kind: 'block',
      value: `${'#'.repeat(Math.min(Math.max(level, 1), 6))} ${value.trim()}`
    }
  }

  const list = listInfo(host, p)
  if (list) {
    const marker = isBullet(host, list) ? '-' : '1.'
    return { kind: 'item', value: `${'    '.repeat(list.level)}${marker} ${value.trim()}` }
  }

  if (isRule(p)) {
    return { kind: 'block', value: '---' }
  }

  return value.trim() ? { kind: 'block', value: escapeStart(value.trim()) } : null
}

function cellMarkdown(host: Host, tc: Element, plain: boolean): string {
  return xml
    .descendants(tc, W, 'p')
    .map((p) => inlineMarkdown(host, p, plain).trim())
    .filter(Boolean)
    .join(' ')
    .replace(/\|/g, '\\|')
}

function tableOf(tr: Element): Node | null {
  let node = tr.parentNode
  while (node && !xml.isElement(node, W, 'tbl')) {
    node = node.parentNode
  }

  return node
}

function tableLine(host: Host, tbl: Element): Line {
  const rows = xml
    .descendants(tbl, W, 'tr')
    .filter((tr) => tableOf(tr) === tbl)
    .map((tr, r) => xml.children(tr, W, 'tc').map((tc) => cellMarkdown(host, tc, r === 0)))
  const columns = Math.max(1, ...rows.map((cells) => cells.length))
  const line = (cells: string[]): string =>
    `| ${Array.from({ length: columns }, (_, c) => cells[c] ?? '').join(' | ')} |`
  const [header = [], ...body] = rows
  const delimiter = `| ${Array.from({ length: columns }, () => '---').join(' | ')} |`

  return { kind: 'block', value: [line(header), delimiter, ...body.map(line)].join('\n') }
}

function blockLine(host: Host, el: Element): Line | null {
  if (el.localName === 'tbl') {
    return tableLine(host, el)
  }

  if (el.localName === 'p') {
    return paragraphLine(host, el)
  }

  const value = escapeStart(text.blockText(el).trim())

  return value ? { kind: 'block', value } : null
}

function range(total: number, opts: MarkdownOptions): [number, number] {
  const from = opts.from ?? 0
  const count = opts.count ?? Math.max(total, 1)
  if (!Number.isInteger(from) || from < 0 || (from >= total && total > 0)) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `from must be a block index from 0 to ${Math.max(0, total - 1)}, got ${String(opts.from)}`,
      { field: 'from', expected: `0 to ${Math.max(0, total - 1)}`, got: opts.from }
    )
  }

  if (!Number.isInteger(count) || count < 1) {
    throw new DocsError('INVALID_OPTIONS', `count must be at least 1, got ${String(opts.count)}`, {
      field: 'count',
      got: opts.count
    })
  }

  return [from, from + count]
}

export function write(host: Host, opts: MarkdownOptions = {}): MarkdownResult {
  const all = blocks(host)
  const [from, to] = range(all.length, opts)
  const lines = all
    .slice(from, to)
    .map((el) => blockLine(host, el))
    .filter((line): line is Line => line !== null)
  let out = ''
  lines.forEach((line, i) => {
    const previous = lines[i - 1]
    if (previous) {
      out += previous.kind === 'item' && line.kind === 'item' ? '\n' : '\n\n'
    }

    out += line.value
  })

  return { total: all.length, markdown: out ? out + '\n' : '' }
}
