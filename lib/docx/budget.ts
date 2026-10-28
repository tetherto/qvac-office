import { DocsError } from '../errors.ts'

export interface Window {
  from: number
  end: number
  total: number
  maxChars?: number
}

export interface Taken<T> {
  items: T[]
  next?: number
}

export function validateMaxChars(value: unknown): number | undefined {
  if (value === undefined || (typeof value === 'number' && Number.isInteger(value) && value > 0)) {
    return value
  }

  throw new DocsError(
    'INVALID_OPTIONS',
    `maxChars must be a positive whole number, got ${String(value)}`,
    {
      field: 'maxChars',
      expected: 'positive integer',
      got: value
    }
  )
}

export function take<T>(window: Window, produce: (index: number) => T): Taken<T> {
  const items: T[] = []
  let used = 0
  for (let i = window.from; i < window.end; i++) {
    const item = produce(i)
    const size = JSON.stringify(item).length
    if (window.maxChars !== undefined && items.length && used + size > window.maxChars) {
      return { items, next: i }
    }

    items.push(item)
    used += size
  }

  return window.end < window.total ? { items, next: window.end } : { items }
}
