import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import { idField } from './ids.ts'
import * as xml from '../xml.ts'
import type { Element } from '../xml.ts'

export interface FindQuery {
  text?: string
  regex?: string
  style?: string
  heading?: number | boolean
  matchCase?: boolean
}

export interface Found {
  index: number
  id?: string
  type: 'paragraph' | 'table' | 'other'
  text: string
  style: string
}

export interface Summary {
  type: Found['type']
  style: string
  styleId: string | null
  heading: number | null
  text: string
}

export type Summarize = (block: Element) => Summary

type Test = (summary: Summary) => boolean

const W = xml.ns.w
const KEYS = ['text', 'regex', 'style', 'heading', 'matchCase']
const PREVIEW = 200

export function blockType(block: Element): Found['type'] {
  if (block.localName === 'p') {
    return 'paragraph'
  }

  return block.localName === 'tbl' ? 'table' : 'other'
}

export function tableStyleId(block: Element): string | null {
  const tblPr = block.localName === 'tbl' ? xml.child(block, W, 'tblPr') : null
  const style = tblPr ? xml.child(tblPr, W, 'tblStyle') : null
  return (style && style.getAttributeNS(W, 'val')) || null
}

function invalid(field: string, expected: string, got: unknown): DocsError {
  return new DocsError('INVALID_OPTIONS', `${field} must be ${expected}`, { field, expected, got })
}

function textTest(value: unknown, matchCase: boolean): Test {
  if (typeof value !== 'string' || value === '') {
    throw invalid('text', 'a non-empty string', value)
  }

  const wanted = matchCase ? value : value.toLowerCase()

  return (s) => (matchCase ? s.text : s.text.toLowerCase()).includes(wanted)
}

function compile(value: string, matchCase: boolean): RegExp {
  try {
    return new RegExp(value, matchCase ? 'u' : 'iu')
  } catch {
    throw invalid('regex', 'a valid regular expression', value)
  }
}

function regexTest(value: unknown, matchCase: boolean): Test {
  if (typeof value !== 'string' || value === '') {
    throw invalid('regex', 'a regular expression source string', value)
  }

  const pattern = compile(value, matchCase)

  return (s) => pattern.test(s.text)
}

function styleTest(value: unknown): Test {
  if (typeof value !== 'string' || value === '') {
    throw invalid('style', 'a style name', value)
  }

  const wanted = value.toLowerCase()

  return (s) => s.style.toLowerCase() === wanted || s.styleId?.toLowerCase() === wanted
}

function headingTest(value: unknown): Test {
  if (typeof value === 'boolean') {
    return (s) => (s.heading !== null) === value
  }

  if (typeof value === 'number' && Number.isInteger(value) && value >= 0 && value <= 9) {
    return (s) => s.heading === value
  }

  throw invalid('heading', 'true, false or a heading level from 0 to 9', value)
}

const BUILDERS: [string, (value: unknown, matchCase: boolean) => Test][] = [
  ['text', textTest],
  ['regex', regexTest],
  ['style', styleTest],
  ['heading', headingTest]
]

function tests(query: Record<string, unknown>): Test[] {
  const unknownKey = Object.keys(query).find((key) => !KEYS.includes(key))
  if (unknownKey) {
    throw invalid('query', `made of ${KEYS.join(', ')}`, unknownKey)
  }

  if (query.matchCase !== undefined && typeof query.matchCase !== 'boolean') {
    throw invalid('matchCase', 'true or false', query.matchCase)
  }

  return BUILDERS.filter(([key]) => query[key] !== undefined).map(([key, build]) =>
    build(query[key], query.matchCase === true)
  )
}

function preview(value: string): string {
  return value.length > PREVIEW ? value.slice(0, PREVIEW - 1) + '\u2026' : value
}

export function find(blocks: Element[], summarize: Summarize, query: unknown): Found[] {
  const checks = isRecord(query) ? tests(query) : []
  if (checks.length === 0) {
    throw new DocsError(
      'INVALID_OPTIONS',
      'find needs at least one of text, regex, style or heading',
      {
        expected: ['text', 'regex', 'style', 'heading'],
        got: query
      }
    )
  }

  const out: Found[] = []
  blocks.forEach((block, index) => {
    const summary = summarize(block)
    if (checks.every((test) => test(summary))) {
      out.push({
        index,
        ...idField(block),
        type: summary.type,
        text: preview(summary.text),
        style: summary.style
      })
    }
  })

  return out
}
