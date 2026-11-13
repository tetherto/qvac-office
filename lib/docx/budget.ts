import { DocsError } from '../errors.ts'

export interface Window {
  from: number
  end: number
  total: number
  limit?: number
}

export interface Taken<T> {
  items: T[]
  next?: number
  truncated?: true
}

export function validateLimit(value: unknown): number | undefined {
  if (value === undefined || (typeof value === 'number' && Number.isInteger(value) && value > 0)) {
    return value
  }

  throw new DocsError(
    'INVALID_OPTIONS',
    `limit is a positive number of characters, got ${String(value)}`,
    {
      field: 'limit',
      expected: 'positive integer',
      got: value
    }
  )
}

export function take<T>(
  window: Window,
  produce: (index: number) => T,
  chars: (item: T) => number
): Taken<T> {
  const items: T[] = []
  let used = 0
  for (let i = window.from; i < window.end; i++) {
    const item = produce(i)
    used += chars(item)
    if (window.limit !== undefined && items.length && used > window.limit) {
      return { items, next: i, truncated: true }
    }

    items.push(item)
  }

  return window.end < window.total ? { items, next: window.end } : { items }
}
