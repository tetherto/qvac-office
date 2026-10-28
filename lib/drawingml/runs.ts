import * as xml from '../xml.ts'
import type { Element } from '../xml.ts'

export interface ReadRun {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  size?: number
  color?: string
  font?: string
  link?: string
  field?: string
}

export type LinkTarget = (rId: string) => string | null

const { a: A, r: R } = xml.ns

function toggle(rPr: Element, name: string): boolean | undefined {
  const value = rPr.getAttribute(name)
  return value === null ? undefined : value === '1' || value === 'true'
}

function linkOf(rPr: Element, target: LinkTarget): string | undefined {
  const hlink = xml.child(rPr, A, 'hlinkClick')
  if (!hlink) {
    return undefined
  }

  const rId = hlink.getAttributeNS(R, 'id')
  const resolved = rId ? target(rId) : null

  return resolved ?? (hlink.getAttribute('action') || undefined)
}

function formatting(rPr: Element, target: LinkTarget): Omit<ReadRun, 'text'> {
  const underline = rPr.getAttribute('u')
  const size = rPr.getAttribute('sz')
  const fill = xml.child(rPr, A, 'solidFill')
  const rgb = fill ? xml.child(fill, A, 'srgbClr') : null
  const latin = xml.child(rPr, A, 'latin')
  const out: Omit<ReadRun, 'text'> = {
    bold: toggle(rPr, 'b'),
    italic: toggle(rPr, 'i'),
    underline: underline === null ? undefined : underline !== 'none',
    size: size === null ? undefined : Number(size) / 100,
    color: rgb?.getAttribute('val')?.toUpperCase() || undefined,
    font: latin?.getAttribute('typeface') || undefined,
    link: linkOf(rPr, target)
  }

  return Object.fromEntries(Object.entries(out).filter(([, value]) => value !== undefined))
}

function runOf(node: Element, target: LinkTarget): ReadRun {
  const rPr = xml.child(node, A, 'rPr')
  const run: ReadRun = { text: xml.text(xml.child(node, A, 't')) }
  if (node.localName === 'fld') {
    run.field = node.getAttribute('type') || 'field'
  }

  return rPr ? { ...run, ...formatting(rPr, target) } : run
}

export function readRuns(p: Element, target: LinkTarget): ReadRun[] {
  const out: ReadRun[] = []
  for (const node of xml.children(p, A)) {
    if (node.localName === 'r' || node.localName === 'fld') {
      out.push(runOf(node, target))
    } else if (node.localName === 'br') {
      out.push({ text: '\n' })
    }
  }

  return out
}
