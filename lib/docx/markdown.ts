export interface InlineFormat {
  bold?: boolean
  italic?: boolean
  code?: boolean
  link?: string
}

export interface Inline extends InlineFormat {
  text: string
}

export type MarkdownBlock =
  | { type: 'heading'; level: number; inlines: Inline[] }
  | { type: 'paragraph'; inlines: Inline[] }
  | { type: 'item'; ordered: boolean; level: number; inlines: Inline[] }
  | { type: 'table'; header: Inline[][]; rows: Inline[][][] }
  | { type: 'rule' }
  | { type: 'code'; lines: string[] }

interface Span {
  inlines: Inline[]
  end: number
}

interface State {
  lines: string[]
  at: number
  out: MarkdownBlock[]
  indents: number[]
}

const ESCAPABLE = /[\\`*_{}[\]()#+\-.!|>~]/
const HEADING = /^ {0,3}(#{1,6})(?:\s+(.*?))?\s*#*\s*$/
const RULE = /^ {0,3}([-*_])(?:\s*\1){2,}\s*$/
const ITEM = /^(\s*)([-*+]|\d{1,9}[.)])\s+(.*)$/
const FENCE = /^ {0,3}(```|~~~)/
const DELIMITER = /^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$/
const FORMAT_KEYS = ['bold', 'italic', 'code', 'link'] as const

function sameFormat(a: Inline, b: Inline): boolean {
  return FORMAT_KEYS.every((key) => a[key] === b[key])
}

function merge(inlines: Inline[]): Inline[] {
  const out: Inline[] = []
  for (const inline of inlines) {
    const last = out[out.length - 1]
    if (last && sameFormat(last, inline)) {
      last.text += inline.text
    } else if (inline.text) {
      out.push({ ...inline })
    }
  }

  return out
}

function codeSpan(src: string, at: number, base: InlineFormat): Span | null {
  const ticks = /^`+/.exec(src.slice(at))?.[0] ?? '`'
  const end = src.indexOf(ticks, at + ticks.length)
  if (end === -1) {
    return null
  }

  let value = src.slice(at + ticks.length, end)
  if (/^ .*[^ ].* $/.test(value)) {
    value = value.slice(1, -1)
  }

  return { inlines: [{ ...base, text: value, code: true }], end: end + ticks.length }
}

function matching(src: string, from: number, pair: string): number {
  let depth = 0
  for (let i = from; i < src.length; i++) {
    const c = src[i]
    if (c === '\\') {
      i++
    } else if (c === pair[0]) {
      depth++
    } else if (c === pair[1]) {
      if (depth === 0) {
        return i
      }

      depth--
    }
  }

  return -1
}

function linkSpan(src: string, at: number, base: InlineFormat): Span | null {
  const close = matching(src, at + 1, '[]')
  if (close === -1 || src[close + 1] !== '(') {
    return null
  }

  const end = matching(src, close + 2, '()')
  if (end === -1) {
    return null
  }

  const target = src
    .slice(close + 2, end)
    .trim()
    .replace(/^<(.*)>$/, '$1')
    .replace(/\s+"[^"]*"$/, '')
  if (!target) {
    return null
  }

  return { inlines: inlines(src.slice(at + 1, close), { ...base, link: target }), end: end + 1 }
}

function runLength(src: string, at: number, c: string): number {
  let n = 0
  while (src[at + n] === c) {
    n++
  }

  return n
}

function closingDelimiter(src: string, from: number, delimiter: string): number {
  const c = delimiter[0]
  for (let j = from; j < src.length; j++) {
    if (src[j] === '\\') {
      j++
      continue
    }

    if (src[j] === '`') {
      const end = codeSpan(src, j, {})
      j = end ? end.end - 1 : j
      continue
    }

    if (src[j] !== c) {
      continue
    }

    const n = runLength(src, j, c)
    const wordAfter = c === '_' && /\w/.test(src[j + n] ?? '')
    if (n === delimiter.length && !/\s/.test(src[j - 1] ?? ' ') && !wordAfter) {
      return j
    }

    j += n - 1
  }

  return -1
}

function emphasisSpan(src: string, at: number, base: InlineFormat): Span | null {
  const c = src[at]
  const n = Math.min(runLength(src, at, c), 3)
  if (c === '_' && /\w/.test(src[at - 1] ?? '')) {
    return null
  }

  if (/\s/.test(src[at + n] ?? ' ')) {
    return null
  }

  const end = closingDelimiter(src, at + n, c.repeat(n))
  if (end === -1) {
    return null
  }

  const format = { ...base }
  if (n !== 1) {
    format.bold = true
  }

  if (n !== 2) {
    format.italic = true
  }

  return { inlines: inlines(src.slice(at + n, end), format), end: end + n }
}

function spanAt(src: string, at: number, base: InlineFormat): Span | null {
  const c = src[at]
  if (c === '`') {
    return codeSpan(src, at, base)
  }

  if (c === '[') {
    return linkSpan(src, at, base)
  }

  if (c === '*' || c === '_') {
    return emphasisSpan(src, at, base)
  }

  return null
}

export function inlines(src: string, base: InlineFormat = {}): Inline[] {
  const out: Inline[] = []
  let buffer = ''
  for (let i = 0; i < src.length;) {
    const span = spanAt(src, i, base)
    if (span) {
      out.push({ ...base, text: buffer }, ...span.inlines)
      buffer = ''
      i = span.end
    } else if (src[i] === '\\' && ESCAPABLE.test(src[i + 1] ?? '')) {
      buffer += src[i + 1]
      i += 2
    } else {
      buffer += src[i]
      i++
    }
  }

  out.push({ ...base, text: buffer })

  return merge(out)
}

function splitRow(line: string): string[] {
  const trimmed = line
    .trim()
    .replace(/^\|/, '')
    .replace(/(?<!\\)\|$/, '')
  return trimmed.split(/(?<!\\)\|/).map((cell) => cell.trim())
}

function isTableStart(lines: string[], at: number): boolean {
  return (
    lines[at].includes('|') && DELIMITER.test(lines[at + 1] ?? '') && lines[at + 1].includes('-')
  )
}

function startsBlock(lines: string[], at: number): boolean {
  const line = lines[at]
  return (
    !line.trim() ||
    HEADING.test(line) ||
    RULE.test(line) ||
    ITEM.test(line) ||
    FENCE.test(line) ||
    isTableStart(lines, at)
  )
}

function continuation(state: State, first: string): string {
  const parts = [first.trim()]
  while (state.at < state.lines.length && !startsBlock(state.lines, state.at)) {
    parts.push(state.lines[state.at].trim())
    state.at++
  }

  return parts.join(' ')
}

function levelOf(state: State, indent: number): number {
  while (state.indents.length && state.indents[state.indents.length - 1] > indent) {
    state.indents.pop()
  }

  const top = state.indents[state.indents.length - 1]
  if (top === undefined || top < indent) {
    state.indents.push(indent)
  }

  return Math.min(state.indents.length - 1, 8)
}

function readTable(state: State): MarkdownBlock {
  const header = splitRow(state.lines[state.at]).map((cell) => inlines(cell))
  const rows: Inline[][][] = []
  state.at += 2
  while (state.at < state.lines.length && state.lines[state.at].includes('|')) {
    rows.push(splitRow(state.lines[state.at]).map((cell) => inlines(cell)))
    state.at++
  }

  return { type: 'table', header, rows }
}

function readFence(state: State, fence: string): MarkdownBlock {
  const lines: string[] = []
  state.at++
  while (state.at < state.lines.length && !state.lines[state.at].trim().startsWith(fence)) {
    lines.push(state.lines[state.at])
    state.at++
  }

  state.at++

  return { type: 'code', lines }
}

function readItem(state: State, match: RegExpExecArray): MarkdownBlock {
  const level = levelOf(state, match[1].replace(/\t/g, '    ').length)
  state.at++
  const body = continuation(state, match[3])

  return { type: 'item', ordered: /\d/.test(match[2]), level, inlines: inlines(body) }
}

function step(state: State): void {
  const line = state.lines[state.at]
  const heading = HEADING.exec(line)
  const item = ITEM.exec(line)
  const fence = FENCE.exec(line)
  if (!line.trim()) {
    state.at++
    return
  }

  if (!item) {
    state.indents = []
  }

  if (fence) {
    state.out.push(readFence(state, fence[1]))
  } else if (heading) {
    state.at++
    state.out.push({
      type: 'heading',
      level: heading[1].length,
      inlines: inlines(heading[2] ?? '')
    })
  } else if (RULE.test(line)) {
    state.at++
    state.out.push({ type: 'rule' })
  } else if (isTableStart(state.lines, state.at)) {
    state.out.push(readTable(state))
  } else if (item) {
    state.out.push(readItem(state, item))
  } else {
    state.at++
    state.out.push({ type: 'paragraph', inlines: inlines(continuation(state, line)) })
  }
}

export function parse(markdown: string): MarkdownBlock[] {
  const state: State = {
    lines: markdown.replace(/\r\n?/g, '\n').split('\n'),
    at: 0,
    out: [],
    indents: []
  }
  while (state.at < state.lines.length) {
    step(state)
  }

  return state.out
}
