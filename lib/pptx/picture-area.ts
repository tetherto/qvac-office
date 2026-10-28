import type { Box } from './inherit.ts'
import type { SlideSize } from './template.ts'

export type Role = 'slot' | 'blank' | 'body' | 'picture' | 'other'

export interface Size {
  readonly width: number
  readonly height: number
}

export interface Placed {
  readonly box: Box
  readonly role: Role
}

export interface Beside<T> {
  readonly shape: T
  readonly box: Box
}

export interface Area<T> {
  readonly box: Box
  readonly cleared: readonly T[]
  readonly beside: Beside<T> | null
}

const FRAME = 0.8
const MARGIN = 0.03
const MIN_BAND = 0.15
const MIN_SPLIT = 0.5

export function framed(slide: SlideSize): Box {
  return {
    x: Math.round((slide.width * (1 - FRAME)) / 2),
    y: Math.round((slide.height * (1 - FRAME)) / 2),
    width: Math.round(slide.width * FRAME),
    height: Math.round(slide.height * FRAME)
  }
}

export function fitted(size: Size, area: Box): Box {
  const scale = Math.min(area.width / size.width, area.height / size.height)
  const width = Math.round(size.width * scale)
  const height = Math.round(size.height * scale)

  return {
    x: area.x + Math.round((area.width - width) / 2),
    y: area.y + Math.round((area.height - height) / 2),
    width,
    height
  }
}

function bandBelow(slide: SlideSize, shapes: readonly Placed[], margin: number): Box | null {
  const top = Math.max(...shapes.map(({ box }) => box.y + box.height)) + margin
  const band = {
    x: margin,
    y: top,
    width: slide.width - 2 * margin,
    height: slide.height - top - margin
  }

  return band.height >= slide.height * MIN_BAND ? band : null
}

function halves(box: Box, gap: number): [Box, Box] {
  if (box.width >= box.height) {
    const half = Math.round((box.width - gap) / 2)
    const rest = box.width - half - gap

    return [
      { ...box, width: half },
      { ...box, x: box.x + half + gap, width: rest }
    ]
  }

  const half = Math.round((box.height - gap) / 2)
  const rest = box.height - half - gap

  return [
    { ...box, height: half },
    { ...box, y: box.y + half + gap, height: rest }
  ]
}

function split<T extends Placed>(shape: T, gap: number): Area<T> {
  const [kept, freed] = halves(shape.box, gap)
  const box = shape.role === 'picture' ? fitted(shape.box, kept) : kept

  return { box: freed, cleared: [], beside: { shape, box } }
}

function largest<T extends Placed>(shapes: readonly T[]): T | undefined {
  return shapes.reduce<T | undefined>(
    (best, shape) =>
      !best || shape.box.width * shape.box.height > best.box.width * best.box.height ? shape : best,
    undefined
  )
}

function overlaps(a: Box, b: Box): boolean {
  return a.x < b.x + b.width && b.x < a.x + a.width && a.y < b.y + b.height && b.y < a.y + a.height
}

function place<T extends Placed>(slide: SlideSize, shapes: readonly T[]): Area<T> {
  const slot = shapes.find((shape) => shape.role === 'slot')
  if (slot) {
    return { box: slot.box, cleared: [slot], beside: null }
  }

  const whole = { box: framed(slide), cleared: [], beside: null }
  if (!shapes.length) {
    return whole
  }

  const margin = Math.round(slide.height * MARGIN)
  const band = bandBelow(slide, shapes, margin)
  if (band) {
    return { box: band, cleared: [], beside: null }
  }

  const bodies = shapes.filter((shape) => shape.role === 'body')
  const wide = bodies.filter((shape) => shape.box.width >= slide.width * MIN_SPLIT)
  // a body already narrowed for a picture keeps its half; the next picture shares the picture's
  const shared =
    largest(wide) ?? largest(shapes.filter((shape) => shape.role === 'picture')) ?? largest(bodies)

  return shared ? split(shared, margin) : whole
}

export function pictureArea<T extends Placed>(slide: SlideSize, shapes: readonly T[]): Area<T> {
  const area = place(
    slide,
    shapes.filter((shape) => shape.role !== 'blank')
  )
  const covered = shapes.filter((shape) => shape.role === 'blank' && overlaps(shape.box, area.box))

  return { ...area, cleared: [...area.cleared, ...covered] }
}
