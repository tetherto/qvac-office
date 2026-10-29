import * as xml from '../xml.ts'
import type { Element } from '../xml.ts'

export interface RunView {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  size?: number
  color?: string
  font?: string
  link?: string
}

export interface RunPlace {
  run: Element
  link: Element | null
  tracked: boolean
  field: boolean
  deleted: boolean
}

export interface Segment {
  node: Element
  place: RunPlace
  start: number
  end: number
}

export type LinkTarget = (hyperlink: Element) => string | undefined

const W = xml.ns.w
const CONTAINERS = new Set([
  'hyperlink',
  'ins',
  'moveTo',
  'smartTag',
  'customXml',
  'fldSimple',
  'dir',
  'bdo',
  'sdtContent',
  'del',
  'moveFrom'
])
const TRACKED_CONTAINERS = new Set(['ins', 'del', 'moveTo', 'moveFrom'])
const FIELD_CONTAINERS = new Set(['fldSimple'])
const DELETED_CONTAINERS = new Set(['del', 'moveFrom'])
const OFF = new Set(['0', 'false', 'off'])

function wVal(el: Element | null): string {
  return (el && el.getAttributeNS(W, 'val')) || ''
}

function nodeText(node: Element): string {
  switch (node.localName) {
    case 't':
      return xml.text(node)
    case 'tab':
      return '\t'
    case 'br':
      return node.getAttributeNS(W, 'type') === 'page' ? '' : '\n'
    case 'cr':
      return '\n'
    case 'noBreakHyphen':
      return '-'
    default:
      return ''
  }
}

interface Walk {
  depth: number
  out: RunPlace[]
}

function hasFormatChange(run: Element): boolean {
  const rPr = xml.child(run, W, 'rPr')
  return !!rPr && !!xml.child(rPr, W, 'rPrChange')
}

function placeOf(run: Element, p: Element, depth: number): RunPlace {
  const place: RunPlace = {
    run,
    link: null,
    tracked: hasFormatChange(run),
    field: depth > 0,
    deleted: false
  }
  for (let el = run.parentNode; el && el !== p; el = el.parentNode) {
    const name = xml.isElement(el, W) ? (el.localName ?? '') : ''
    if (name === 'hyperlink' && xml.isElement(el)) {
      place.link = place.link ?? el
    }

    place.tracked = place.tracked || TRACKED_CONTAINERS.has(name)
    place.field = place.field || FIELD_CONTAINERS.has(name)
    place.deleted = place.deleted || DELETED_CONTAINERS.has(name)
  }

  return place
}

function fieldStep(run: Element, walk: Walk): void {
  for (const mark of xml.children(run, W, 'fldChar')) {
    const type = mark.getAttributeNS(W, 'fldCharType')
    if (type === 'begin') {
      walk.depth++
    } else if (type === 'end') {
      walk.depth = Math.max(0, walk.depth - 1)
    }
  }
}

function collect(parent: Element, p: Element, walk: Walk): void {
  for (const el of xml.children(parent, W)) {
    const name = el.localName ?? ''
    if (name === 'r') {
      const depth = walk.depth
      fieldStep(el, walk)
      walk.out.push(placeOf(el, p, Math.max(depth, walk.depth)))
    } else if (name === 'sdt') {
      collect(xml.child(el, W, 'sdtContent') ?? el, p, walk)
    } else if (CONTAINERS.has(name)) {
      collect(el, p, walk)
    }
  }
}

export function places(p: Element): RunPlace[] {
  const walk: Walk = { depth: 0, out: [] }
  collect(p, p, walk)
  return walk.out
}

export function segments(list: RunPlace[]): Segment[] {
  const out: Segment[] = []
  let at = 0
  for (const place of list.filter((item) => !item.deleted)) {
    for (const node of xml.children(place.run, W)) {
      const length = nodeText(node).length
      if (length > 0) {
        out.push({ node, place, start: at, end: at + length })
        at += length
      }
    }
  }

  return out
}

function toggle(rPr: Element | null, local: string): boolean {
  const el = rPr ? xml.child(rPr, W, local) : null
  return !!el && !OFF.has(wVal(el))
}

function child(rPr: Element | null, local: string): Element | null {
  return rPr ? xml.child(rPr, W, local) : null
}

function fontOf(rPr: Element | null): string | undefined {
  const fonts = child(rPr, 'rFonts')
  return (
    (fonts && (fonts.getAttributeNS(W, 'ascii') || fonts.getAttributeNS(W, 'hAnsi'))) || undefined
  )
}

function formatOf(rPr: Element | null): Omit<RunView, 'text'> {
  const size = Number(wVal(child(rPr, 'sz'))) / 2
  const color = wVal(child(rPr, 'color'))
  const underline = child(rPr, 'u')
  const found: Omit<RunView, 'text'> = {
    bold: toggle(rPr, 'b'),
    italic: toggle(rPr, 'i'),
    underline: !!underline && wVal(underline) !== 'none',
    size: size > 0 ? size : undefined,
    color: color && color !== 'auto' ? color : undefined,
    font: fontOf(rPr)
  }
  for (const key of Object.keys(found)) {
    const value: unknown = Reflect.get(found, key)
    if (value === undefined || value === false) {
      Reflect.deleteProperty(found, key)
    }
  }

  return found
}

function viewOf(place: RunPlace, target: LinkTarget): RunView {
  const text = xml.children(place.run, W).map(nodeText).join('')
  const link = place.link ? target(place.link) : undefined
  const view: RunView = { text, ...formatOf(xml.child(place.run, W, 'rPr')) }
  if (link) {
    view.link = link
  }

  return view
}

function sameFormat(a: RunView, b: RunView): boolean {
  return JSON.stringify({ ...a, text: '' }) === JSON.stringify({ ...b, text: '' })
}

export function readRuns(p: Element, target: LinkTarget): RunView[] {
  const out: RunView[] = []
  const live = places(p).filter((place) => !place.deleted)
  for (const view of live.map((place) => viewOf(place, target))) {
    if (!view.text) continue
    const last = out[out.length - 1]
    if (last && sameFormat(last, view)) {
      last.text += view.text
    } else {
      out.push(view)
    }
  }

  return out
}
