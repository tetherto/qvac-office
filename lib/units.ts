import { DocsError } from './errors.ts'

const EMU_PER_INCH = 914400
const EMU_PER_POINT = 12700
const EMU_PER_CM = 360000

const EMU_PER_UNIT: Readonly<Record<string, number>> = {
  emu: 1,
  in: EMU_PER_INCH,
  cm: EMU_PER_CM,
  mm: EMU_PER_CM / 10,
  pt: EMU_PER_POINT,
  px: 9525
}

const LENGTH = /^\s*(-?\d+(?:\.\d+)?|-?\.\d+)\s*(emu|in|cm|mm|pt|px)\s*$/i

export type Length = number | string

export interface LengthBox {
  x?: Length
  y?: Length
  width?: Length
  height?: Length
}

export interface LengthSize {
  width?: Length
  height?: Length
}

export interface EmuBox {
  x: number | undefined
  y: number | undefined
  width: number | undefined
  height: number | undefined
}

export interface EmuSize {
  width: number | undefined
  height: number | undefined
}

export function inches(n: number): number {
  return Math.round(n * EMU_PER_INCH)
}

export function points(n: number): number {
  return Math.round(n * EMU_PER_POINT)
}

export function cm(n: number): number {
  return Math.round(n * EMU_PER_CM)
}

function malformed(key: string, value: unknown): DocsError {
  return new DocsError(
    'INVALID_BOX',
    `${key} must be EMU or a length such as "2in", "5cm", "40mm", "72pt" or "96px", got ${JSON.stringify(value)}`,
    { field: key, got: value, expected: 'integer EMU or "<number><in|cm|mm|pt|px|emu>"' }
  )
}

export function length(value: unknown, key = 'length'): number | undefined {
  if (value === undefined || typeof value === 'number') {
    return value
  }

  const match = typeof value === 'string' ? LENGTH.exec(value) : null
  if (!match?.[1] || !match[2]) {
    throw malformed(key, value)
  }

  return Math.round(Number(match[1]) * (EMU_PER_UNIT[match[2].toLowerCase()] ?? 1))
}

export function box<T extends LengthBox>(opts: T): T & EmuBox {
  return {
    ...opts,
    x: length(opts.x, 'x'),
    y: length(opts.y, 'y'),
    width: length(opts.width, 'width'),
    height: length(opts.height, 'height')
  }
}

export function size<T extends LengthSize>(opts: T): T & EmuSize {
  return { ...opts, width: length(opts.width, 'width'), height: length(opts.height, 'height') }
}

export type WithLengths<T, K extends keyof T> = Omit<T, K> & { [P in K]?: Length }

export type BoxKey = 'x' | 'y' | 'width' | 'height'
export type SizeKey = 'width' | 'height'
