import { Document } from './document.ts'
import type { OpenOptions } from '../opc/package.ts'
import type { CreateOptions } from './template.ts'

export { Document } from './document.ts'
export type * from './document.ts'

export function open(bytes: Uint8Array | ArrayBuffer, opts?: OpenOptions): Document {
  return Document.open(bytes, opts)
}

export function create(opts?: CreateOptions): Document {
  return Document.create(opts)
}
