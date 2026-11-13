import * as inherit from './inherit.ts'
import { DocsError } from '../errors.ts'
import { shapesOf, type ShapeEntry } from './shapes.ts'
import { isEmptyPlaceholder } from './overlaps.ts'
import type { Package } from '../opc/package.ts'
import type { Box } from './inherit.ts'
import type { Placement } from './types.ts'
import type { IntoOptions } from './layout-types.ts'

export interface Claim {
  id: number
  box: Box
}

const BOX_KEYS = ['x', 'y', 'width', 'height'] as const

function placeholders(pkg: Package, part: string): string[] {
  return shapesOf(pkg.xml(part))
    .filter((s) => s.placeholder !== null)
    .map((s) => `${s.id} (${s.placeholder})`)
}

export function checkInto(opts: Placement & IntoOptions): void {
  const given = BOX_KEYS.filter((key) => opts[key] !== undefined)
  if (opts.into !== undefined && given.length) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `into takes the placeholder's box; leave out ${given.join(', ')}`,
      { field: 'into', got: given }
    )
  }
}

export interface ClaimScope {
  pkg: Package
  part: string
}

export function claim(scope: ClaimScope, found: ShapeEntry, replace: boolean): Claim {
  const { pkg, part } = scope
  if (found.placeholder === null || found.id === null) {
    const candidates = placeholders(pkg, part)
    throw new DocsError(
      'NOT_PLACEHOLDER',
      `shape ${found.id} is not a placeholder; the slide's placeholders are: ${candidates.join(', ') || 'none'}`,
      { field: 'into', got: found.id, candidates }
    )
  }

  if (!replace && !isEmptyPlaceholder(found)) {
    throw new DocsError(
      'PLACEHOLDER_HAS_TEXT',
      `placeholder ${found.id} (${found.placeholder}) is not empty; pass replace: true to put the new shape in its place`,
      { field: 'replace', expected: true }
    )
  }

  const box = inherit.box({ pkg, slide: part, el: found.el })
  if (!box) {
    throw new DocsError('INVALID_BOX', `placeholder ${found.id} has no box on its slide or layout`)
  }

  return { id: found.id, box }
}
