import { Document } from './document.ts'
import type { OpenOptions } from '../opc/package.ts'
import type { CreateOptions } from './template.ts'

export { Document } from './document.ts'
export type * from './document.ts'
export type { BlockRef } from './ids.ts'
export type { FindQuery, Found } from './find.ts'
export type { Outline, OutlineHeading, OutlineOptions } from './outline.ts'
export type { RunView } from './runs.ts'
export type { TextQuery } from './match.ts'
export type { FormatTextResult, LinkResult, ReplaceOptions, ReplaceResult } from './ranges.ts'
export type { RunFormat } from './text.ts'

export function open(bytes: Uint8Array | ArrayBuffer, opts?: OpenOptions): Document {
  return Document.open(bytes, opts)
}

export function create(opts?: CreateOptions): Document {
  return Document.create(opts)
}
