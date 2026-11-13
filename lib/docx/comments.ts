import * as xml from '../xml.ts'
import * as text from './text.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import { contentStart, locate } from './anchor.ts'
import {
  W,
  W14,
  W15,
  MC,
  REL_BASE,
  WML,
  declare,
  existingPart,
  freeParaId,
  nextId,
  paraIds,
  part,
  requireText,
  setParaId,
  styleId,
  touchMain,
  type Host
} from './host.ts'
import type { Element } from '../xml.ts'

export interface CommentOptions {
  text: string
  author?: string
  initials?: string
  find?: string
}

export interface CommentAdded {
  id: string
}

export interface CommentDeleted {
  id: string
  anchors: number
}

export interface ResolveOptions {
  done?: boolean
}

export interface CommentResolved {
  id: string
  done: boolean
  changed: boolean
}

const COMMENTS_REL = REL_BASE + 'comments'
const EXTENDED_REL = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended'
const IDS_REL = 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds'
const ANCHORS = ['commentRangeStart', 'commentRangeEnd', 'commentReference']

function commentsPart(host: Host): string {
  return part(host, {
    rel: COMMENTS_REL,
    type: WML + 'comments+xml',
    name: '/word/comments.xml',
    root: `<w:comments xmlns:w="${W}" xmlns:w14="${W14}" xmlns:mc="${MC}" mc:Ignorable="w14"/>`
  })
}

function extendedPart(host: Host): string {
  return part(host, {
    rel: EXTENDED_REL,
    type: WML + 'commentsExtended+xml',
    name: '/word/commentsExtended.xml',
    root: `<w15:commentsEx xmlns:w15="${W15}" xmlns:mc="${MC}" mc:Ignorable="w15"/>`
  })
}

function options(input: unknown): CommentOptions {
  if (!isRecord(input)) {
    throw new DocsError('INVALID_OPTIONS', 'give the comment as { text, author, initials, find }')
  }

  const out: CommentOptions = { text: requireText(input.text, 'text') }
  for (const key of ['author', 'initials', 'find'] as const) {
    if (input[key] !== undefined) {
      out[key] = requireText(input[key], key)
    }
  }

  return out
}

function body(host: Host, comment: Element, value: string): void {
  const doc = xml.owner(comment)
  const taken = paraIds(host)
  const style = styleId(host, 'annotation text') ?? styleId(host, 'CommentText')
  value.split('\n').forEach((line, i) => {
    const p = xml.append(comment, text.wEl(doc, 'p'))
    setParaId(p, freeParaId(taken))
    if (style) {
      text.ensure(p, 'pPr').appendChild(text.wEl(doc, 'pStyle', { val: style }))
    }

    if (i === 0) {
      const mark = xml.append(p, referenceRun(host, doc))
      mark.appendChild(text.wEl(doc, 'annotationRef'))
    }

    p.appendChild(text.buildParagraph(doc, line, null).firstChild ?? text.wEl(doc, 'r'))
  })
}

function referenceRun(host: Host, doc: ReturnType<typeof xml.owner>): Element {
  const r = text.wEl(doc, 'r')
  const style =
    styleId(host, 'annotation reference', 'character') ??
    styleId(host, 'CommentReference', 'character')
  if (style) {
    xml.append(r, text.wEl(doc, 'rPr')).appendChild(text.wEl(doc, 'rStyle', { val: style }))
  }

  return r
}

interface Placement {
  opts: CommentOptions
  id: string
}

function anchor(host: Host, p: Element, { opts, id }: Placement): void {
  const doc = xml.owner(p)
  const start = text.wEl(doc, 'commentRangeStart', { id })
  const end = text.wEl(doc, 'commentRangeEnd', { id })
  const reference = referenceRun(host, doc)
  reference.appendChild(text.wEl(doc, 'commentReference', { id }))
  if (opts.find === undefined) {
    p.insertBefore(start, contentStart(p))
    p.appendChild(end)
    p.appendChild(reference)

    return
  }

  const span = locate(p, opts.find, 'the paragraph')
  span.first.parentNode?.insertBefore(start, span.first)
  span.last.parentNode?.insertBefore(end, span.last.nextSibling)
  end.parentNode?.insertBefore(reference, end.nextSibling)
}

export function add(host: Host, p: Element, input: unknown): CommentAdded {
  const opts = options(input)
  const name = commentsPart(host)
  const root = host.package.root(name)
  const id = String(nextId(xml.children(root, W, 'comment')))
  anchor(host, p, { opts, id })
  const comment = xml.append(
    root,
    text.wEl(xml.owner(root), 'comment', {
      id,
      author: opts.author ?? 'QVAC',
      initials: opts.initials
    })
  )
  body(host, comment, opts.text)
  host.package.touch(name)
  touchMain(host)

  return { id }
}

interface Found {
  part: string
  comment: Element
}

function find(host: Host, input: unknown): Found {
  const id = String(input)
  const name = existingPart(host, COMMENTS_REL)
  const all = name ? xml.children(host.package.root(name), W, 'comment') : []
  const comment = all.find((el) => text.wGet(el, 'id') === id)
  if (!name || !comment) {
    const ids = all.map((el) => text.wGet(el, 'id'))
    throw new DocsError(
      'COMMENT_NOT_FOUND',
      `no comment with id ${id}; the comments are ${ids.join(', ') || 'none'}`,
      { field: 'id', got: input, candidates: ids }
    )
  }

  return { part: name, comment }
}

function ownParaIds(comment: Element): string[] {
  return xml.children(comment, W, 'p').flatMap((p) => p.getAttributeNS(W14, 'paraId') || [])
}

function dropExtensions(host: Host, ids: string[]): void {
  for (const rel of [EXTENDED_REL, IDS_REL]) {
    const name = existingPart(host, rel)
    if (!name) {
      continue
    }

    const stale = xml
      .children(host.package.root(name))
      .filter((el) =>
        Array.from(el.attributes).some((a) => a.localName === 'paraId' && ids.includes(a.value))
      )
    for (const el of stale) {
      xml.remove(el)
    }

    if (stale.length) {
      host.package.touch(name)
    }
  }
}

export function remove(host: Host, input: unknown): CommentDeleted {
  const { part: name, comment } = find(host, input)
  const id = text.wGet(comment, 'id')
  const main = host.package.root(host.main)
  const anchors = ANCHORS.flatMap((local) => xml.descendants(main, W, local)).filter(
    (el) => text.wGet(el, 'id') === id
  )
  for (const el of anchors) {
    const run = el.localName === 'commentReference' ? el.parentNode : null
    xml.remove(el)
    if (run && xml.children(run, W).every((n) => n.localName === 'rPr')) {
      xml.remove(run)
    }
  }

  dropExtensions(host, ownParaIds(comment))
  xml.remove(comment)
  host.package.touch(name)
  touchMain(host)

  return { id, anchors: anchors.length }
}

function lastParaId(host: Host, comment: Element, name: string): string {
  const paragraphs = xml.children(comment, W, 'p')
  const last =
    paragraphs[paragraphs.length - 1] ?? xml.append(comment, text.wEl(xml.owner(comment), 'p'))
  const own = last.getAttributeNS(W14, 'paraId')
  if (own) {
    return own
  }

  const id = freeParaId(paraIds(host))
  setParaId(last, id)
  host.package.touch(name)

  return id
}

export function resolve(host: Host, input: unknown, opts: ResolveOptions = {}): CommentResolved {
  const done = opts.done ?? true
  if (typeof done !== 'boolean') {
    throw new DocsError('INVALID_OPTIONS', 'done must be true or false', {
      field: 'done',
      got: done
    })
  }

  const { part: name, comment } = find(host, input)
  const id = text.wGet(comment, 'id')
  const paraId = lastParaId(host, comment, name)
  const extended = extendedPart(host)
  const root = host.package.root(extended)
  declare(root, 'w15', W15)
  const entry =
    xml
      .children(root, W15, 'commentEx')
      .find((el) => el.getAttributeNS(W15, 'paraId') === paraId) ??
    xml.append(root, xml.builder(xml.owner(root), W15)('commentEx'))
  const value = done ? '1' : '0'
  const changed = entry.getAttributeNS(W15, 'done') !== value
  entry.setAttributeNS(W15, 'w15:paraId', paraId)
  entry.setAttributeNS(W15, 'w15:done', value)
  host.package.touch(extended)

  return { id, done, changed }
}
