import { Presentation } from './presentation.ts'
import type { OpenOptions } from '../opc/package.ts'
import type { CreateOptions } from './template.ts'

export { Presentation } from './presentation.ts'
export type * from './types.ts'
export type * from './layout-types.ts'
export type { CreateOptions, SlideSize, ThemeColor, ThemeFonts } from './template.ts'
export type { FindOptions, FoundShape, SlideOutline } from './find.ts'
export type { PageOptions, SlidePage } from './page.ts'
export type {
  ReplaceTextOptions,
  ReplacedText,
  FormattedText,
  LinkedText,
  TextShape
} from './text-ranges.ts'
export type { Search } from '../drawingml/ranges.ts'
export type { ReadRun } from '../drawingml/runs.ts'

export function open(bytes: Uint8Array, opts?: OpenOptions): Presentation {
  return Presentation.open(bytes, opts)
}

export function create(opts?: CreateOptions): Presentation {
  return Presentation.create(opts)
}
