import * as xml from '../xml.ts'
import * as text from './text.ts'
import { DocsError } from '../errors.ts'
import type { Document, Element } from '../xml.ts'
import type { Package } from '../opc/package.ts'

export interface StyleInfo {
  el: Element
  id: string
  type: string
  name: string
  isDefault: boolean
  basedOn: string | null
}

export interface FormattedStyle {
  style: string
}

const W = xml.ns.w

const STYLE_ORDER = [
  'name',
  'aliases',
  'basedOn',
  'next',
  'link',
  'autoRedefine',
  'hidden',
  'uiPriority',
  'semiHidden',
  'unhideWhenUsed',
  'qFormat',
  'locked',
  'personal',
  'personalCompose',
  'personalReply',
  'rsid',
  'pPr',
  'rPr',
  'tblPr',
  'trPr',
  'tcPr',
  'tblStylePr'
]

function displayName(name: string): string {
  if (name !== name.toLowerCase()) return name
  return name.replace(/(^|\s)\S/g, (c) => c.toUpperCase())
}

function wVal(el: Element | null): string | null {
  return el ? el.getAttributeNS(W, 'val') : null
}

export class Styles {
  package: Package
  part: string | null

  constructor(pkg: Package, part: string | null) {
    this.package = pkg
    this.part = part
  }

  get doc(): Document | null {
    return this.part ? this.package.xml(this.part) : null
  }

  all(): StyleInfo[] {
    const doc = this.doc
    if (!doc) {
      return []
    }

    return xml.children(xml.root(doc), W, 'style').map((el) => {
      const id = text.wGet(el, 'styleId')
      const name = wVal(xml.child(el, W, 'name')) || id
      return {
        el,
        id,
        type: text.wGet(el, 'type') || 'paragraph',
        name: displayName(name),
        isDefault: ['1', 'true', 'on'].includes(text.wGet(el, 'default')),
        basedOn: wVal(xml.child(el, W, 'basedOn'))
      }
    })
  }

  names(type = 'paragraph'): string[] {
    return this.all()
      .filter((s) => s.type === type)
      .map((s) => s.name)
  }

  byId(id: string | null): StyleInfo | null {
    return this.all().find((s) => s.id === id) ?? null
  }

  resolve(ref: unknown, type = 'paragraph'): StyleInfo | null {
    const wanted = String(ref).toLowerCase()
    const styles = this.all().filter((s) => s.type === type)
    return (
      styles.find((s) => s.id === ref) ??
      styles.find((s) => s.name.toLowerCase() === wanted) ??
      styles.find((s) => s.id.toLowerCase() === wanted.replace(/\s+/g, '')) ??
      null
    )
  }

  require(ref: unknown, type = 'paragraph'): StyleInfo {
    const style = this.resolve(ref, type)
    if (!style) {
      const names = this.names(type).map((n) => `"${n}"`)
      throw new DocsError(
        'STYLE_NOT_FOUND',
        `the document has no ${type} style "${String(ref)}"; it has ${names.join(', ') || 'none'}`
      )
    }

    return style
  }

  defaultParagraph(): StyleInfo | null {
    return this.all().find((s) => s.type === 'paragraph' && s.isDefault) ?? null
  }

  nameOf(styleId: string | null): string {
    const style = styleId ? this.byId(styleId) : this.defaultParagraph()
    if (style) {
      return style.name
    }

    return styleId || 'Normal'
  }

  headingLevel(styleId: string | null): number | null {
    let style = styleId ? this.byId(styleId) : this.defaultParagraph()
    const seen = new Set<string>()
    while (style && !seen.has(style.id)) {
      seen.add(style.id)
      const level = ownHeadingLevel(style)
      if (level !== null) {
        return level
      }

      style = style.basedOn ? this.byId(style.basedOn) : null
    }

    return null
  }

  numIdOf(styleId: string | null): string | null {
    for (const style of this.chain(styleId)) {
      const pPr = xml.child(style.el, W, 'pPr')
      const numPr = pPr ? xml.child(pPr, W, 'numPr') : null
      const numId = numPr ? wVal(xml.child(numPr, W, 'numId')) : null
      if (numId) {
        return numId
      }
    }

    return null
  }

  isList(styleId: string | null): boolean {
    const numId = this.numIdOf(styleId)
    return numId !== null && numId !== '0'
  }

  format(ref: unknown, input: unknown): FormattedStyle {
    const format = text.validateFormat(input)
    const style = this.require(ref)
    const doc = xml.owner(style.el)
    const { align, ...run } = format
    if (Object.keys(run).length) {
      const rPr =
        xml.child(style.el, W, 'rPr') ??
        xml.insertOrdered(style.el, text.wEl(doc, 'rPr'), STYLE_ORDER)
      text.applyRunFormat(rPr, run)
    }

    if (align) {
      const pPr =
        xml.child(style.el, W, 'pPr') ??
        xml.insertOrdered(style.el, text.wEl(doc, 'pPr'), STYLE_ORDER)
      text.wAttr(text.ensure(pPr, 'jc', text.PARAGRAPH_ORDER), 'val', text.ALIGNMENTS[align])
    }

    if (this.part) {
      this.package.touch(this.part)
    }

    return { style: style.name }
  }

  private chain(styleId: string | null): StyleInfo[] {
    const out: StyleInfo[] = []
    const seen = new Set<string>()
    let style = styleId ? this.byId(styleId) : null
    while (style && !seen.has(style.id)) {
      seen.add(style.id)
      out.push(style)
      style = style.basedOn ? this.byId(style.basedOn) : null
    }

    return out
  }
}

function ownHeadingLevel(style: StyleInfo): number | null {
  const match = /^heading ([1-9])$/i.exec(style.name)
  if (match) {
    return Number(match[1])
  }

  if (style.name.toLowerCase() === 'title') {
    return 0
  }

  const pPr = xml.child(style.el, W, 'pPr')
  const outline = pPr ? wVal(xml.child(pPr, W, 'outlineLvl')) : null
  if (outline !== null && outline !== '' && Number(outline) < 9) {
    return Number(outline) + 1
  }

  return null
}
