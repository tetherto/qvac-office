import * as xml from '../xml.ts'
import * as text from '../drawingml/text.ts'
import { shapesOf, type ShapeEntry } from './shapes.ts'
import { slideBox, intersects } from './geometry.ts'
import type { Element } from '../xml.ts'
import type { Package } from '../opc/package.ts'
import type { SlideSize } from './template.ts'

export interface OverlapScope {
  pkg: Package
  slide: string
  size: SlideSize
}

const { p: P } = xml.ns
const CHROME = new Set(['dt', 'ftr', 'sldNum'])
const BACKGROUND_SHARE = 0.9

function topLevel(el: Element): Element {
  let top = el
  for (let up = el.parentNode; xml.isElement(up, P, 'grpSp'); up = up.parentNode) {
    top = up
  }

  return top
}

export function isEmptyPlaceholder(shape: ShapeEntry): boolean {
  if (shape.placeholder === null || shape.el.localName !== 'sp') {
    return false
  }

  const txBody = xml.child(shape.el, P, 'txBody')

  return !txBody || !text.readParagraphs(txBody).some((p) => p.text.trim())
}

function isHidden(shape: ShapeEntry): boolean {
  const nv = xml.children(shape.el)[0]
  const hidden = nv ? xml.child(nv, P, 'cNvPr')?.getAttribute('hidden') : null
  return hidden === '1' || hidden === 'true'
}

function counts(shape: ShapeEntry): boolean {
  return !(
    shape.group !== null ||
    shape.id === null ||
    CHROME.has(shape.placeholder ?? '') ||
    isEmptyPlaceholder(shape) ||
    isHidden(shape)
  )
}

function isBackground(scope: OverlapScope, shape: ShapeEntry, area: number): boolean {
  return shape.kind === 'picture' && area >= scope.size.width * scope.size.height * BACKGROUND_SHARE
}

export function overlaps(scope: OverlapScope, el: Element): number[] {
  const own = slideBox({ pkg: scope.pkg, slide: scope.slide, el })
  if (!own) {
    return []
  }

  const self = topLevel(el)
  const out: number[] = []
  for (const shape of shapesOf(scope.pkg.xml(scope.slide))) {
    if (shape.el === self || !counts(shape)) continue
    const box = slideBox({ pkg: scope.pkg, slide: scope.slide, el: shape.el })
    if (box && !isBackground(scope, shape, box.width * box.height) && intersects(own, box)) {
      out.push(shape.id ?? 0)
    }
  }

  return out
}
