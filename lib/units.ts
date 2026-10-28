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

export function inches(n: number): number {
  return Math.round(n * EMU_PER_INCH)
}

export function points(n: number): number {
  return Math.round(n * EMU_PER_POINT)
}

export function cm(n: number): number {
  return Math.round(n * EMU_PER_CM)
}

export function length(value: unknown, key = 'length'): unknown {
  if (typeof value !== 'string') {
    return value
  }

  const match = LENGTH.exec(value)
  if (!match?.[1] || !match[2]) {
    throw new DocsError(
      'INVALID_BOX',
      `${key} must be EMU or a length such as "2in", "5cm", "40mm", "72pt" or "96px", got "${value}"`,
      { field: key, got: value, expected: 'integer EMU or "<number><in|cm|mm|pt|px|emu>"' }
    )
  }

  return Math.round(Number(match[1]) * (EMU_PER_UNIT[match[2].toLowerCase()] ?? 1))
}

export function lengths<T>(opts: T, keys: readonly string[]): T {
  if (opts === null || typeof opts !== 'object') {
    return opts
  }

  const out: Record<string, unknown> = { ...(opts as Record<string, unknown>) }
  for (const key of keys) {
    if (key in out) {
      out[key] = length(out[key], key)
    }
  }

  return out as T
}

export type Length = number | string

export type WithLengths<T, K extends keyof T> = Omit<T, K> & { [P in K]?: Length }

export const BOX_KEYS = ['x', 'y', 'width', 'height'] as const
export const SIZE_KEYS = ['width', 'height'] as const

export type BoxKey = (typeof BOX_KEYS)[number]
export type SizeKey = (typeof SIZE_KEYS)[number]
