import type { Box } from './inherit.ts'
import type { ShapeKind } from './shapes.ts'
import type { ReadParagraph, Losses, TextInput, Style } from '../drawingml/text.ts'
import type { ChartInfo, ChartSpec } from '../drawingml/chart.ts'

export type { Box } from './inherit.ts'
export type { ShapeKind } from './shapes.ts'

export type ShapeId = number | 'title' | 'subtitle' | 'body'

export interface ShapeRef {
  slide: number
  shape: ShapeId
}

export interface CellRef extends ShapeRef {
  row: number
  column: number
}

export interface ReadOptions {
  from?: number
  count?: number
  slides?: number[]
}

export interface ShapeInfo {
  id: number | null
  name: string
  kind: ShapeKind
  placeholder?: string
  group?: number
  paragraphs?: ReadParagraph[]
  rows?: string[][]
  box?: Box
  chart?: ChartInfo
}

export interface SlideInfo {
  index: number
  layout: string | null
  shapes: ShapeInfo[]
  notes?: string
}

export interface LayoutInfo {
  name: string
  type: string
}

export interface TextResult {
  paragraphs: number
  changed: boolean
  removed?: Losses
  overflow?: boolean
  fontScale?: number
}

export interface FitOptions {
  fit?: 'shrink'
}

export interface StyleOptions {
  paragraph?: number
}

export interface RunCount {
  runs: number
}

export interface AddSlideOptions {
  like?: number
  layout?: string
  at?: number
  title?: TextInput
  body?: TextInput
  image?: Uint8Array
}

export interface AddedSlide {
  index: number
  layout: string
  shapes: ShapeInfo[]
}

export interface Placement {
  x?: number
  y?: number
  width?: number
  height?: number
}

export interface PictureOptions extends Placement {
  name?: string
  description?: string
}

export interface TextBoxOptions extends Placement {
  text?: TextInput
  style?: Style
  fill?: string
  name?: string
}

export interface TableOptions extends Placement {
  rows: (string | number | null)[][]
  header?: boolean
  name?: string
}

export interface ChartOptions extends Partial<ChartSpec>, Placement {
  name?: string
}

export interface PlacedShape extends Box {
  id: number
}

export interface ReplaceOptions {
  stretch?: boolean
}

export interface ReplacedPicture extends Partial<Box> {
  id: number | null
}

export interface MoveOptions {
  to: number
}

export interface Moved {
  moved: boolean
}

export interface Deleted {
  deleted: number | null
  removed?: { animations: true }
}

export interface Remaining {
  slides: number
}

export interface Changed {
  changed: boolean
}
