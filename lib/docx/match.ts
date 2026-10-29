import * as xml from '../xml.ts'
import { paragraphText } from './text.ts'
import { places, segments, type RunPlace, type Segment } from './runs.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { DocsErrorCode } from '../codes.ts'
import type { Element } from '../xml.ts'

export interface TextQuery {
  find: string
  all?: boolean
  matchCase?: boolean
}

export type Range = [number, number]
export type Rule = 'replace' | 'format' | 'link' | 'delete'
export type Target = [number, Element]

export interface Planned {
  index: number
  p: Element
  ranges: Range[]
}

interface Refusal {
  code: DocsErrorCode
  reason: string
}

interface Checked {
  index: number
  rule: Rule
}

const W = xml.ns.w
const FLAGS = ['all', 'matchCase'] as const

function invalid(field: string, expected: string, got: unknown): DocsError {
  return new DocsError('INVALID_OPTIONS', `${field} must be ${expected}`, { field, expected, got })
}

export function validateQuery(input: unknown, field = 'query'): TextQuery {
  if (!isRecord(input) || typeof input.find !== 'string' || input.find === '') {
    throw invalid(
      `${field}.find`,
      'the non-empty text to look for',
      isRecord(input) ? input.find : input
    )
  }

  const flag = FLAGS.find((key) => input[key] !== undefined && typeof input[key] !== 'boolean')
  if (flag) {
    throw invalid(`${field}.${flag}`, 'true or false', input[flag])
  }

  return { find: input.find, all: input.all === true, matchCase: input.matchCase === true }
}

function escape(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

export function matchRanges(value: string, query: TextQuery): Range[] {
  const pattern = new RegExp(escape(query.find), query.matchCase ? 'g' : 'gi')
  const found: Range[] = Array.from(value.matchAll(pattern), (m) => [
    m.index,
    m.index + m[0].length
  ])

  return query.all ? found : found.slice(0, 1)
}

function inTextBox(p: Element, block: Element): boolean {
  for (let node = p.parentNode; node && node !== block; node = node.parentNode) {
    if (xml.isElement(node, W, 'txbxContent')) {
      return true
    }
  }

  return false
}

export function paragraphsOf(block: Element): Element[] {
  if (block.localName === 'p') {
    return [block]
  }

  return xml.descendants(block, W, 'p').filter((p) => !inTextBox(p, block))
}

function planBlock([index, block]: Target, query: TextQuery): Planned[] {
  return paragraphsOf(block)
    .map((p) => ({ index, p, ranges: matchRanges(paragraphText(p), query) }))
    .filter((item) => item.ranges.length > 0)
}

export function plan(targets: Target[], query: TextQuery): Planned[] {
  const out: Planned[] = []
  for (const target of targets) {
    out.push(...planBlock(target, query))
    if (!query.all && out.length) {
      return out.slice(0, 1)
    }
  }

  return out
}

export function overlapping(segs: Segment[], [start, end]: Range): Segment[] {
  return segs.filter((seg) => seg.start < end && seg.end > start)
}

function spanned(list: RunPlace[], touched: Segment[]): RunPlace[] {
  const first = list.indexOf(touched[0].place)
  const last = list.indexOf(touched[touched.length - 1].place)
  return list.slice(first, last + 1)
}

function refusal(span: RunPlace[], rule: Rule): Refusal | null {
  if (span.some((place) => place.tracked)) {
    return {
      code: 'TRACKED_CHANGES',
      reason: 'has tracked changes; accept or reject them in Word first'
    }
  }

  if (rule !== 'format' && span.some((place) => place.field)) {
    return {
      code: 'FIELD_TEXT',
      reason: 'is the result of a field, which Word recalculates; edit the field instead'
    }
  }

  if (rule === 'link' && span.some((place) => place.link)) {
    return { code: 'ALREADY_LINKED', reason: 'is already inside a hyperlink' }
  }

  const bounded = rule === 'replace' || rule === 'link'
  if (bounded && new Set(span.map((place) => place.run.parentNode)).size > 1) {
    return {
      code: 'RANGE_BOUNDARY',
      reason: 'crosses the edge of a hyperlink or content control; match text on one side of it'
    }
  }

  return null
}

function check(p: Element, range: Range, where: Checked): void {
  const list = places(p)
  const found = refusal(spanned(list, overlapping(segments(list), range)), where.rule)
  if (!found) {
    return
  }

  const matched = paragraphText(p).slice(range[0], range[1])

  throw new DocsError(found.code, `"${matched}" in block ${where.index} ${found.reason}`, {
    block: where.index,
    text: matched
  })
}

export function checkAll(planned: Planned[], rule: Rule): void {
  for (const { index, p, ranges } of planned) {
    for (const range of ranges) {
      check(p, range, { index, rule })
    }
  }
}

export function count(planned: Planned[]): number {
  return planned.reduce((sum, item) => sum + item.ranges.length, 0)
}
