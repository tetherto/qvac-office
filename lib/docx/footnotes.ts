import * as xml from '../xml.ts'
import * as text from './text.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import { locate } from './anchor.ts'
import {
  W,
  R,
  REL_BASE,
  WML,
  nextId,
  part,
  requireText,
  styleId,
  touchMain,
  type Host
} from './host.ts'
import type { Document, Element } from '../xml.ts'

export interface FootnoteOptions {
  text: string
  find?: string
}

export interface FootnoteAdded {
  id: string
}

const SEPARATORS = `<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>`

function footnotesPart(host: Host): string {
  return part(host, {
    rel: REL_BASE + 'footnotes',
    type: WML + 'footnotes+xml',
    name: '/word/footnotes.xml',
    root: `<w:footnotes xmlns:w="${W}" xmlns:r="${R}">${SEPARATORS}</w:footnotes>`
  })
}

function options(input: unknown): FootnoteOptions {
  if (!isRecord(input)) {
    throw new DocsError('INVALID_OPTIONS', 'give the footnote as { text, find }')
  }

  const out: FootnoteOptions = { text: requireText(input.text, 'text') }
  if (input.find !== undefined) {
    out.find = requireText(input.find, 'find')
  }

  return out
}

interface Mark {
  local: string
  id?: string
}

interface NoteSpec {
  id: string
  value: string
}

function referenceRun(host: Host, doc: Document, mark: Mark): Element {
  const r = text.wEl(doc, 'r')
  const rPr = xml.append(r, text.wEl(doc, 'rPr'))
  const style = styleId(host, 'footnote reference', 'character')
  if (style) {
    rPr.appendChild(text.wEl(doc, 'rStyle', { val: style }))
  } else {
    rPr.appendChild(text.wEl(doc, 'vertAlign', { val: 'superscript' }))
  }

  r.appendChild(text.wEl(doc, mark.local, { id: mark.id }))

  return r
}

function note(host: Host, doc: Document, opts: NoteSpec): Element {
  const footnote = text.wEl(doc, 'footnote', { id: opts.id })
  const style = styleId(host, 'footnote text')
  opts.value.split('\n').forEach((line, i) => {
    const p = xml.append(footnote, text.wEl(doc, 'p'))
    if (style) {
      text.ensure(p, 'pPr').appendChild(text.wEl(doc, 'pStyle', { val: style }))
    }

    if (i === 0) {
      p.appendChild(referenceRun(host, doc, { local: 'footnoteRef' }))
    }

    const run = text.buildParagraph(doc, i === 0 ? ` ${line}` : line, null).firstChild
    if (run) {
      p.appendChild(run)
    }
  })

  return footnote
}

export function add(host: Host, p: Element, input: unknown): FootnoteAdded {
  const opts = options(input)
  const name = footnotesPart(host)
  const root = host.package.root(name)
  const id = String(Math.max(1, nextId(xml.children(root, W, 'footnote'))))
  const reference = referenceRun(host, xml.owner(p), { local: 'footnoteReference', id })
  if (opts.find === undefined) {
    p.appendChild(reference)
  } else {
    const span = locate(p, opts.find, 'the paragraph')
    span.last.parentNode?.insertBefore(reference, span.last.nextSibling)
  }

  root.appendChild(note(host, xml.owner(root), { id, value: opts.text }))
  host.package.touch(name)
  touchMain(host)

  return { id }
}
