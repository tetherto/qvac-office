import type { Box } from './inherit.ts'
import type { ThemeColor } from './template.ts'
import type { ShapeId, SlideRef } from './types.ts'
import type { Length } from '../units.ts'

export type SizePreset = '16:9' | '4:3'

export interface Size {
  width: number
  height: number
  preset?: SizePreset
}

export interface SizeInput {
  width: Length
  height: Length
}

export interface SizeOptions {
  scale?: boolean
}

export interface Resized extends Size {
  scaled: number
}

export interface IntoOptions {
  into?: ShapeId
  replace?: boolean
}

export interface PlacedResult extends Box {
  id: number
  overlaps: number[]
  replaced?: number
}

export type LineDash =
  | 'solid'
  | 'dot'
  | 'dash'
  | 'lgDash'
  | 'dashDot'
  | 'lgDashDot'
  | 'lgDashDotDot'
  | 'sysDash'
  | 'sysDot'
  | 'sysDashDot'
  | 'sysDashDotDot'

export interface Line {
  color?: string
  width?: Length
  dash?: LineDash
}

export interface DuplicateOptions {
  slide?: SlideRef
  x?: Length
  y?: Length
}

export interface Duplicated extends PlacedResult {
  slide: number
}

export type ZOrder = 'front' | 'back' | 'forward' | 'backward'

export interface Reordered {
  moved: boolean
  position: number
}

export interface GroupOptions {
  name?: string
}

export interface Grouped extends Box {
  id: number
  children: number[]
}

export interface Ungrouped {
  ungrouped: number
  children: number[]
  removed?: { animations: true }
}

export interface Theme {
  name: string
  colors: Record<ThemeColor, string>
  fonts: { heading: string; body: string }
}

export interface ThemeInput {
  colors?: Partial<Record<ThemeColor, string>>
  fonts?: { heading?: string; body?: string }
}

export interface ThemeChanged {
  changed: string[]
  parts: number
}

export interface CarriedPlaceholder {
  id: number
  from: string
  to: string
}

export interface KeptPlaceholder {
  id: number
  placeholder: string
  text: string
}

export interface Relayout {
  layout: string
  carried: CarriedPlaceholder[]
  kept: KeptPlaceholder[]
  removed: number[]
  added: number[]
  removedAnimations?: true
}
