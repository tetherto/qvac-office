import { Presentation } from './presentation.ts'
import type { OpenOptions } from '../opc/package.ts'
import type { CreateOptions } from './template.ts'

export { Presentation } from './presentation.ts'
export type * from './types.ts'
export type { CreateOptions, SlideSize, ThemeColor, ThemeFonts } from './template.ts'

export function open(bytes: Uint8Array, opts?: OpenOptions): Presentation {
  return Presentation.open(bytes, opts)
}

export function create(opts?: CreateOptions): Presentation {
  return Presentation.create(opts)
}
