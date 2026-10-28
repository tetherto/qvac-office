import * as xml from '../xml.ts'
import * as units from '../units.ts'
import * as slides from './slides.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import { SIZES, resolveSize, type SlideSize } from './template.ts'
import type { Element } from '../xml.ts'
import type { Size, SizePreset, Resized } from './layout-types.ts'

interface Factors {
  x: number
  y: number
}

const { a: A, p: P } = xml.ns
const DEFAULT_SIZE: SlideSize = { width: 9144000, height: 6858000 }
const HORIZONTAL = new Set(['x', 'cx', 'w'])

function presetOf(size: SlideSize): SizePreset | undefined {
  if (size.width * 9 === size.height * 16) return '16:9'
  if (size.width * 3 === size.height * 4) return '4:3'
  return undefined
}

export function read(deck: slides.Deck): Size {
  const sz = xml.child(deck.pkg.root(deck.main), P, 'sldSz')
  const size = sz
    ? { width: Number(sz.getAttribute('cx')), height: Number(sz.getAttribute('cy')) }
    : DEFAULT_SIZE
  const preset = presetOf(size)

  return preset ? { ...size, preset } : size
}

function resolve(input: unknown): SlideSize {
  if (!(typeof input === 'string' && input in SIZES) && !isRecord(input)) {
    throw new DocsError('INVALID_OPTIONS', 'size must be "16:9", "4:3" or { width, height }', {
      field: 'size',
      expected: ['16:9', '4:3', '{ width, height }'],
      got: input
    })
  }

  if (typeof input === 'string') {
    return resolveSize(input)
  }

  return resolveSize({
    width: units.length(input.width, 'width'),
    height: units.length(input.height, 'height')
  })
}

function scaleAttributes(el: Element, by: Factors): void {
  for (const name of ['x', 'y', 'cx', 'cy']) {
    const value = el.getAttribute(name)
    if (value !== null) {
      const factor = HORIZONTAL.has(name) ? by.x : by.y
      el.setAttribute(name, String(Math.round(Number(value) * factor)))
    }
  }
}

function scaleTree(root: Element, by: Factors): number {
  let count = 0
  const frames = [...xml.descendants(root, A, 'xfrm'), ...xml.descendants(root, P, 'xfrm')]
  for (const frame of frames) {
    for (const part of xml.children(frame, A)) {
      scaleAttributes(part, by)
    }

    count++
  }

  for (const col of xml.descendants(root, A, 'gridCol')) {
    col.setAttribute('w', String(Math.round(Number(col.getAttribute('w')) * by.x)))
  }

  for (const row of xml.descendants(root, A, 'tr')) {
    row.setAttribute('h', String(Math.round(Number(row.getAttribute('h')) * by.y)))
  }

  return count
}

function scaledParts(deck: slides.Deck): string[] {
  const layouts = slides.allLayouts(deck)
  const parts = [
    ...slides.entries(deck).map((e) => e.part),
    ...layouts.map((l) => l.part),
    ...layouts.map((l) => l.master)
  ]

  return [...new Set(parts)]
}

function scaleDeck(deck: slides.Deck, by: Factors): number {
  let count = 0
  for (const part of scaledParts(deck)) {
    const tree = xml.descendants(deck.pkg.xml(part), P, 'spTree')[0]
    const scaled = tree ? scaleTree(tree, by) : 0
    if (scaled) {
      deck.pkg.touch(part)
      count += scaled
    }
  }

  return count
}

function writeSize(deck: slides.Deck, size: SlideSize): void {
  const top = deck.pkg.root(deck.main)
  const sz =
    xml.child(top, P, 'sldSz') ??
    xml.insertOrdered(top, xml.builder(xml.owner(top), P)('sldSz'), slides.PRESENTATION_ORDER)
  xml.setAttributes(sz, { cx: size.width, cy: size.height })
  if (presetOf(size) === '4:3' && size.width === SIZES['4:3'].width) {
    sz.setAttribute('type', 'screen4x3')
  } else {
    sz.removeAttribute('type')
  }

  deck.pkg.touch(deck.main)
}

export function write(deck: slides.Deck, input: unknown, scale: boolean): Resized {
  const next = resolve(input)
  const before = read(deck)
  writeSize(deck, next)
  const by = { x: next.width / before.width, y: next.height / before.height }
  const scaled = scale && (by.x !== 1 || by.y !== 1) ? scaleDeck(deck, by) : 0

  return { ...read(deck), scaled }
}
