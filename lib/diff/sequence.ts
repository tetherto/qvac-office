export interface Entry<T> {
  index: number
  item: T
}

export interface Edit<T> {
  kind: 'same' | 'added' | 'removed' | 'changed'
  before?: Entry<T>
  after?: Entry<T>
}

interface Sides<T> {
  before: readonly Entry<T>[]
  after: readonly Entry<T>[]
  a: readonly string[]
  b: readonly string[]
}

function lengths(a: readonly string[], b: readonly string[]): number[][] {
  const table = Array.from({ length: a.length + 1 }, () => new Array<number>(b.length + 1).fill(0))
  for (let i = a.length - 1; i >= 0; i--) {
    for (let j = b.length - 1; j >= 0; j--) {
      const longest =
        a[i] === b[j]
          ? (table[i + 1]?.[j + 1] ?? 0) + 1
          : Math.max(table[i + 1]?.[j] ?? 0, table[i]?.[j + 1] ?? 0)
      table[i]?.splice(j, 1, longest)
    }
  }

  return table
}

function pairGap<T>(gap: Edit<T>[]): Edit<T>[] {
  const removed = gap.filter((edit) => edit.kind === 'removed')
  const added = gap.filter((edit) => edit.kind === 'added')
  const paired = Math.min(removed.length, added.length)
  const changed = removed
    .slice(0, paired)
    .map((edit, i): Edit<T> => ({ kind: 'changed', before: edit.before, after: added[i]?.after }))

  return [...changed, ...removed.slice(paired), ...added.slice(paired)]
}

function pairGaps<T>(edits: Edit<T>[]): Edit<T>[] {
  const out: Edit<T>[] = []
  let gap: Edit<T>[] = []
  for (const edit of edits) {
    if (edit.kind === 'same') {
      out.push(...pairGap(gap), edit)
      gap = []
    } else {
      gap.push(edit)
    }
  }

  return [...out, ...pairGap(gap)]
}

function walk<T>({ before, after, a, b }: Sides<T>): Edit<T>[] {
  const table = lengths(a, b)
  const edits: Edit<T>[] = []
  let i = 0
  let j = 0
  while (i < before.length || j < after.length) {
    const left = before[i]
    const right = after[j]
    if (left && right && a[i] === b[j]) {
      edits.push({ kind: 'same', before: left, after: right })
      i++
      j++
    } else if (right && (!left || (table[i]?.[j + 1] ?? 0) >= (table[i + 1]?.[j] ?? 0))) {
      edits.push({ kind: 'added', after: right })
      j++
    } else {
      edits.push({ kind: 'removed', before: left })
      i++
    }
  }

  return edits
}

export function sequence<T>(
  before: readonly T[],
  after: readonly T[],
  key: (item: T) => string
): Edit<T>[] {
  return pairGaps(
    walk({
      before: before.map((item, index) => ({ index, item })),
      after: after.map((item, index) => ({ index, item })),
      a: before.map(key),
      b: after.map(key)
    })
  )
}
