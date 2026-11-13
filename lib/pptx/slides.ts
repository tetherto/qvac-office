import * as xml from '../xml.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Element } from '../xml.ts'
import type { Package } from '../opc/package.ts'
import { layoutName, layoutType, hasBody } from './read.ts'
import type { AddSlideOptions } from './types.ts'

export interface Deck {
  pkg: Package
  main: string
}

export interface SlideEntry {
  el: Element
  id: string
  rId: string
  part: string
}

interface LayoutEntry {
  master: string
  part: string
}

const { p: P, r: R, p14: P14 } = xml.ns
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
export const PRESENTATION_ORDER = [
  'sldMasterIdLst',
  'notesMasterIdLst',
  'handoutMasterIdLst',
  'sldIdLst',
  'sldSz'
]

export function slideIdList(deck: Deck): Element | null {
  return xml.child(deck.pkg.root(deck.main), P, 'sldIdLst')
}

export function ensureSlideIdList(deck: Deck): Element {
  const existing = slideIdList(deck)
  if (existing) {
    return existing
  }

  const top = deck.pkg.root(deck.main)

  return xml.insertOrdered(top, xml.builder(xml.owner(top), P)('sldIdLst'), PRESENTATION_ORDER)
}

export function entries(deck: Deck): SlideEntry[] {
  const list = slideIdList(deck)
  if (!list) {
    return []
  }

  const rels = deck.pkg.rels(deck.main)

  return xml.children(list, P, 'sldId').map((el) => {
    const rId = el.getAttributeNS(R, 'id')
    const rel = rels.get(rId)
    if (!rel || rel.external || !rel.part || !deck.pkg.has(rel.part)) {
      throw new DocsError(
        'BROKEN_PPTX',
        `slide id ${String(el.getAttribute('id'))} points at a missing part`
      )
    }

    return { el, id: xml.attr(el, 'id'), rId: rId ?? '', part: rel.part }
  })
}

export function slideLabel(ref: unknown): string {
  return isRecord(ref) && 'id' in ref ? `id ${String(ref.id)}` : String(ref)
}

function slideNotFound(all: SlideEntry[], ref: unknown): DocsError {
  const range = all.length ? ` (0 to ${all.length - 1})` : ''
  const ids = all.map((e) => Number(e.id))
  return new DocsError(
    'SLIDE_NOT_FOUND',
    `slide ${slideLabel(ref)} does not exist; the deck has ${all.length} slides${range} with ids ${ids.join(', ') || 'none'}`,
    { field: 'slide', got: ref, expected: 'a slide index or { id }', candidates: ids }
  )
}

function positionOf(all: SlideEntry[], ref: unknown): number {
  if (isRecord(ref) && 'id' in ref) {
    return Number.isInteger(ref.id) ? all.findIndex((e) => e.id === String(ref.id)) : -1
  }

  return typeof ref === 'number' && Number.isInteger(ref) && ref >= 0 && ref < all.length ? ref : -1
}

export function slideIndex(deck: Deck, ref: unknown): number {
  const all = entries(deck)
  const index = positionOf(all, ref)
  if (index === -1) {
    throw slideNotFound(all, ref)
  }

  return index
}

export function slideAt(deck: Deck, ref: unknown): SlideEntry {
  return entries(deck)[slideIndex(deck, ref)]
}

export function layoutOf(pkg: Package, slidePart: string): string | null {
  return pkg.related(slidePart, REL_BASE + 'slideLayout')[0] ?? null
}

export function allLayouts(deck: Deck): LayoutEntry[] {
  return deck.pkg
    .related(deck.main, REL_BASE + 'slideMaster')
    .flatMap((master) =>
      deck.pkg.related(master, REL_BASE + 'slideLayout').map((part) => ({ master, part }))
    )
}

function layoutByName(deck: Deck, layouts: LayoutEntry[], name: string): string {
  const wanted = name.toLowerCase()
  const match = layouts.find(({ part }) => layoutName(deck.pkg.xml(part)).toLowerCase() === wanted)
  if (match) {
    return match.part
  }

  const names = layouts.map(({ part }) => `"${layoutName(deck.pkg.xml(part))}"`)

  throw new DocsError(
    'LAYOUT_NOT_FOUND',
    `no layout is named "${name}"; the deck has ${names.join(', ')}`
  )
}

function defaultLayout(deck: Deck, layouts: LayoutEntry[]): string {
  const { pkg } = deck
  const used = entries(deck)
    .map((e) => layoutOf(pkg, e.part))
    .filter(
      (part): part is string =>
        part !== null && layoutType(pkg.xml(part)) !== 'title' && hasBody(pkg.xml(part))
    )
  const last = used[used.length - 1]
  if (last) {
    return last
  }

  const content = layouts.find(({ part }) => ['obj', 'tx'].includes(layoutType(pkg.xml(part))))

  return (content ?? layouts[0]).part
}

export function pickLayout(deck: Deck, opts: AddSlideOptions): string {
  if (opts.like !== undefined && opts.layout !== undefined) {
    throw new DocsError('INVALID_OPTIONS', 'pass either like or layout, not both')
  }

  if (opts.like !== undefined) {
    const layout = layoutOf(deck.pkg, slideAt(deck, opts.like).part)
    if (!layout) {
      throw new DocsError('BROKEN_PPTX', `slide ${opts.like} has no layout`)
    }

    return layout
  }

  const layouts = allLayouts(deck)
  if (layouts.length === 0) {
    throw new DocsError('BROKEN_PPTX', 'the deck has no slide layouts')
  }

  if (opts.layout !== undefined) {
    return layoutByName(deck, layouts, String(opts.layout))
  }

  return defaultLayout(deck, layouts)
}

function sectionIds(list: Element): string[] {
  return xml.children(list, P14, 'sldId').map((ref) => xml.attr(ref, 'id'))
}

export function placeInSection(deck: Deck, id: string): void {
  const lists = xml.descendants(deck.pkg.root(deck.main), P14, 'sldIdLst')
  if (lists.length === 0) {
    return
  }

  for (const list of lists) {
    for (const ref of xml.children(list, P14, 'sldId')) {
      if (ref.getAttribute('id') === id) {
        xml.remove(ref)
      }
    }
  }

  const order = entries(deck).map((e) => e.id)
  const pos = order.indexOf(id)
  const neighbor = order[pos - 1] ?? order[pos + 1]
  const home = lists.find((list) => sectionIds(list).includes(neighbor)) ?? lists[0]
  home.appendChild(
    xml.setAttributes(xml.named(xml.owner(home), `${home.prefix || 'p14'}:sldId`, P14), { id })
  )
  const sorted = xml
    .children(home, P14, 'sldId')
    .sort((x, y) => order.indexOf(xml.attr(x, 'id')) - order.indexOf(xml.attr(y, 'id')))
  for (const el of sorted) {
    home.appendChild(el)
  }
}
