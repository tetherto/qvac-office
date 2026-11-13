import * as xml from '../xml.ts'
import * as text from './text.ts'
import { W, SECTION_ORDER } from './host.ts'
import type { Element } from '../xml.ts'

export interface ChangeCounts {
  insertions: number
  deletions: number
  moves: number
  formatting: number
  paragraphs: number
  rows: number
}

type Mode = 'accept' | 'reject'

const MOVE_RANGES = ['moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd']
const PROPERTY_CHANGES = [
  'rPrChange',
  'pPrChange',
  'sectPrChange',
  'tblPrChange',
  'trPrChange',
  'tcPrChange',
  'tblPrExChange',
  'tblGridChange',
  'numberingChange'
]
const KEPT_ON_RESTORE: Record<string, readonly string[]> = {
  pPr: ['rPr', 'sectPr'],
  sectPr: ['headerReference', 'footerReference']
}
const ORDERS: Record<string, readonly string[]> = {
  pPr: text.PARAGRAPH_ORDER,
  sectPr: SECTION_ORDER
}

function zero(): ChangeCounts {
  return { insertions: 0, deletions: 0, moves: 0, formatting: 0, paragraphs: 0, rows: 0 }
}

function unwrap(el: Element): void {
  const parent = el.parentNode
  if (!parent) {
    return
  }

  for (const child of xml.children(el)) {
    parent.insertBefore(child, el)
  }

  xml.remove(el)
}

function isParagraphMark(el: Element): boolean {
  const rPr = el.parentNode
  return xml.isElement(rPr, W, 'rPr') && xml.isElement(rPr.parentNode, W, 'pPr')
}

function isRowMark(el: Element): boolean {
  return xml.isElement(el.parentNode, W, 'trPr')
}

function marks(root: Element, local: string, test: (el: Element) => boolean): Element[] {
  return xml.descendants(root, W, local).filter(test)
}

function paragraphOf(mark: Element): Element | null {
  const p = mark.parentNode?.parentNode?.parentNode ?? null
  return xml.isElement(p, W, 'p') ? p : null
}

function joinNext(p: Element): void {
  let next = p.nextSibling
  while (next && !xml.isElement(next)) {
    next = next.nextSibling
  }

  if (!xml.isElement(next, W, 'p')) {
    return
  }

  const pPr = xml.child(next, W, 'pPr')
  const at = pPr ? pPr.nextSibling : next.firstChild
  for (const child of xml.children(p)) {
    if (child.localName !== 'pPr') {
      next.insertBefore(child, at)
    }
  }

  xml.remove(p)
}

function paragraphMarks(root: Element, mode: Mode, counts: ChangeCounts): void {
  const joining = mode === 'accept' ? ['del', 'moveFrom'] : ['ins', 'moveTo']
  for (const local of ['ins', 'del', 'moveFrom', 'moveTo']) {
    for (const mark of marks(root, local, isParagraphMark)) {
      const p = joining.includes(local) ? paragraphOf(mark) : null
      xml.remove(mark)
      if (p) {
        joinNext(p)
      }

      counts.paragraphs++
    }
  }
}

function rowMarks(root: Element, mode: Mode, counts: ChangeCounts): void {
  const dropping = mode === 'accept' ? 'del' : 'ins'
  for (const local of ['ins', 'del']) {
    for (const mark of marks(root, local, isRowMark)) {
      const row = mark.parentNode?.parentNode ?? null
      xml.remove(mark)
      if (local === dropping && row) {
        xml.remove(row)
      }

      counts.rows++
    }
  }
}

function restoreText(el: Element): void {
  const doc = xml.owner(el)
  for (const [from, to] of [
    ['delText', 't'],
    ['delInstrText', 'instrText']
  ]) {
    for (const old of xml.descendants(el, W, from)) {
      const replacement = text.wEl(doc, to)
      for (const attr of Array.from(old.attributes)) {
        replacement.setAttributeNS(attr.namespaceURI, attr.name, attr.value)
      }

      for (const child of Array.from(old.childNodes)) {
        replacement.appendChild(child)
      }

      old.parentNode?.replaceChild(replacement, old)
    }
  }
}

function restore(change: Element): void {
  const parent = change.parentNode
  if (!xml.isElement(parent)) {
    return
  }

  const local = parent.localName ?? ''
  const kept = KEPT_ON_RESTORE[local] ?? []
  for (const el of xml.children(parent)) {
    if (el !== change && !kept.includes(el.localName ?? '')) {
      xml.remove(el)
    }
  }

  const old = xml.child(change)
  const order = ORDERS[local]
  for (const el of old ? xml.children(old) : []) {
    if (order) {
      xml.insertOrdered(parent, el, order)
    } else {
      parent.insertBefore(el, change)
    }
  }

  xml.remove(change)
}

function content(root: Element, mode: Mode, counts: ChangeCounts): void {
  for (const local of MOVE_RANGES) {
    for (const el of xml.descendants(root, W, local)) {
      xml.remove(el)
    }
  }

  const kept = mode === 'accept' ? ['ins', 'moveTo'] : ['del', 'moveFrom']
  const dropped = mode === 'accept' ? ['del', 'moveFrom'] : ['ins', 'moveTo']
  for (const local of dropped) {
    for (const el of xml.descendants(root, W, local)) {
      xml.remove(el)
      bump(counts, local)
    }
  }

  for (const local of kept) {
    for (const el of xml.descendants(root, W, local)) {
      restoreText(el)
      unwrap(el)
      bump(counts, local)
    }
  }
}

function bump(counts: ChangeCounts, local: string): void {
  if (local === 'ins') {
    counts.insertions++
  } else if (local === 'del') {
    counts.deletions++
  } else {
    counts.moves++
  }
}

function properties(root: Element, mode: Mode, counts: ChangeCounts): void {
  for (const local of PROPERTY_CHANGES) {
    for (const el of xml.descendants(root, W, local)) {
      if (mode === 'reject') {
        restore(el)
      } else {
        xml.remove(el)
      }

      counts.formatting++
    }
  }
}

export function resolve(roots: Element[], mode: Mode): ChangeCounts {
  const counts = zero()
  for (const root of roots) {
    paragraphMarks(root, mode, counts)
    rowMarks(root, mode, counts)
    content(root, mode, counts)
    properties(root, mode, counts)
  }

  return counts
}

export function total(counts: ChangeCounts): number {
  return Object.values(counts).reduce((sum, n) => sum + n, 0)
}
