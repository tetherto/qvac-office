import { DocsError } from '../errors.ts'
import { idField } from './ids.ts'
import { blockText } from './text.ts'
import type { Element } from '../xml.ts'

export interface OutlineOptions {
  depth?: number
}

export interface OutlineHeading {
  level: number
  title: string
  index: number
  id?: string
  end: number
}

export interface Outline {
  total: number
  chars: number
  headings: OutlineHeading[]
}

export interface Heading {
  index: number
  level: number
  title: string
  end: number
}

const TITLE_LENGTH = 120

function depthOf(opts: OutlineOptions): number {
  const depth = opts.depth ?? 9
  if (!Number.isInteger(depth) || depth < 1 || depth > 9) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `depth must be a heading level from 1 to 9, got ${String(opts.depth)}`,
      {
        field: 'depth',
        expected: 'integer from 1 to 9',
        got: opts.depth
      }
    )
  }

  return depth
}

function short(value: string): string {
  return value.length > TITLE_LENGTH ? value.slice(0, TITLE_LENGTH - 1) + '…' : value
}

export function outline(
  blocks: Element[],
  headings: Heading[],
  opts: OutlineOptions = {}
): Outline {
  const depth = depthOf(opts)
  return {
    total: blocks.length,
    chars: blocks.reduce((sum, block) => sum + blockText(block).length, 0),
    headings: headings
      .filter((heading) => heading.level <= depth)
      .map((heading) => ({
        level: heading.level,
        title: short(heading.title),
        index: heading.index,
        ...idField(blocks[heading.index]),
        end: heading.end
      }))
  }
}
