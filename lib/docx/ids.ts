import * as xml from '../xml.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import type { Element } from '../xml.ts'

export type BlockRef = number | { id: string }

export interface IdRef {
  id: string
}

export const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const XMLNS = 'http://www.w3.org/2000/xmlns/'
const W = xml.ns.w
const LIMIT = 0x80000000
const HEX_ID = /^[0-9A-Fa-f]{8}$/

export function paraId(p: Element): string | null {
  return (p.getAttributeNS(W14, 'paraId') ?? '').toUpperCase() || null
}

export function blockId(block: Element): string | null {
  const p = block.localName === 'p' ? block : xml.descendants(block, W, 'p')[0]
  return p ? paraId(p) : null
}

export function keepId(block: Element, id: string | null): void {
  const paragraphs = xml.descendants(block, W, 'p')
  const first = paragraphs[0]
  if (!id || !first || paraId(first) === id) {
    return
  }

  const name = { namespace: W14, local: 'paraId' }
  const displaced = first.getAttributeNS(W14, 'paraId')
  const holder = paragraphs.find((p) => paraId(p) === id)
  if (holder && displaced) {
    xml.setNsAttribute(holder, name, displaced)
  } else if (holder) {
    holder.removeAttributeNS(W14, 'paraId')
  }

  xml.setNsAttribute(first, name, id)
}

function usedIds(root: Element): Set<string> {
  const used = new Set<string>()
  for (const p of xml.descendants(root, W, 'p')) {
    for (const name of ['paraId', 'textId']) {
      const value = p.getAttributeNS(W14, name)
      if (value) {
        used.add(value.toUpperCase())
      }
    }
  }

  return used
}

function fresh(used: Set<string>): string {
  for (;;) {
    const id = Math.floor(Math.random() * LIMIT)
      .toString(16)
      .toUpperCase()
      .padStart(8, '0')
    if (id !== '00000000' && !used.has(id)) {
      used.add(id)
      return id
    }
  }
}

function declare(root: Element, namespace: string, prefix: string): string {
  const found = root.lookupPrefix(namespace)
  if (found) {
    return found
  }

  root.setAttributeNS(XMLNS, `xmlns:${prefix}`, namespace)

  return prefix
}

function markIgnorable(root: Element, prefix: string): void {
  const mc = declare(root, MC, 'mc')
  const listed = (root.getAttributeNS(MC, 'Ignorable') ?? '').split(/\s+/).filter(Boolean)
  if (!listed.includes(prefix)) {
    root.setAttributeNS(MC, `${mc}:Ignorable`, [...listed, prefix].join(' '))
  }
}

export function stamp(root: Element, blocks: Element[]): string[] {
  const prefix = declare(root, W14, 'w14')
  markIgnorable(root, prefix)
  const used = usedIds(root)
  for (const block of blocks) {
    const paragraphs = block.localName === 'p' ? [block] : xml.descendants(block, W, 'p')
    for (const p of paragraphs) {
      p.setAttributeNS(W14, `${prefix}:paraId`, fresh(used))
      p.setAttributeNS(W14, `${prefix}:textId`, fresh(used))
    }
  }

  return blocks.map((block) => blockId(block) ?? '')
}

function containing(blocks: Element[], id: string): number {
  return blocks.findIndex((block) => xml.descendants(block, W, 'p').some((p) => paraId(p) === id))
}

function checkedId(value: unknown): string {
  const id = typeof value === 'string' && HEX_ID.test(value) ? value.toUpperCase() : null
  if (!id) {
    throw new DocsError(
      'BLOCK_NOT_FOUND',
      `a block id is 8 hex digits (w14:paraId), got ${JSON.stringify(value)}`,
      { field: 'id', expected: '8 hex digits', got: value }
    )
  }

  return id
}

export function resolve(blocks: Element[], ref: BlockRef): number {
  if (!isRecord(ref)) {
    return ref
  }

  const id = checkedId(ref.id)
  const index = blocks.findIndex((block) => blockId(block) === id)
  if (index !== -1) {
    return index
  }

  const inside = containing(blocks, id)
  if (inside !== -1) {
    throw new DocsError(
      'BLOCK_NOT_FOUND',
      `${id} is a paragraph inside block ${inside}, not a block of its own; use block ${inside}`,
      { field: 'id', got: id, inside }
    )
  }

  throw new DocsError(
    'BLOCK_NOT_FOUND',
    `no block has id ${id}; read() or find() lists the current ids`,
    { field: 'id', got: id }
  )
}

export interface Located {
  index: number
  p: Element
}

export function locate(blocks: Element[], ref: IdRef): Located {
  const id = checkedId(ref.id)
  for (const [index, block] of blocks.entries()) {
    const list = block.localName === 'p' ? [block] : xml.descendants(block, W, 'p')
    const p = list.find((el) => paraId(el) === id)
    if (p) {
      return { index, p }
    }
  }

  throw new DocsError(
    'BLOCK_NOT_FOUND',
    `no paragraph has id ${id}; read() or find() lists the current ids`,
    { field: 'id', got: id }
  )
}

export interface IdField {
  id?: string
}

export function idField(block: Element): IdField {
  const out: IdField = {}
  const id = blockId(block)
  if (id) {
    out.id = id
  }

  return out
}
