export interface Ordered {
  order: number[]
  cyclic: Set<number>
}

interface Walk {
  edges: number[][]
  index: Int32Array
  low: Int32Array
  onStack: Uint8Array
  stack: number[]
  counter: number
  out: Ordered
}

function emit(walk: Walk, root: number): void {
  const members: number[] = []
  let node = walk.stack.pop()
  while (node !== undefined) {
    walk.onStack[node] = 0
    members.push(node)
    node = node === root ? undefined : walk.stack.pop()
  }

  const cyclic = members.length > 1 || walk.edges[root].includes(root)
  for (const member of members) {
    walk.out.order.push(member)
    if (cyclic) {
      walk.out.cyclic.add(member)
    }
  }
}

function visit(walk: Walk, start: number): void {
  const frames: [number, number][] = [[start, 0]]
  walk.index[start] = walk.low[start] = walk.counter++
  walk.stack.push(start)
  walk.onStack[start] = 1
  while (frames.length) {
    const frame = frames[frames.length - 1]
    const [node, next] = frame
    const targets = walk.edges[node]
    if (next < targets.length) {
      frame[1]++
      const target = targets[next]
      if (walk.index[target] === -1) {
        walk.index[target] = walk.low[target] = walk.counter++
        walk.stack.push(target)
        walk.onStack[target] = 1
        frames.push([target, 0])
      } else if (walk.onStack[target]) {
        walk.low[node] = Math.min(walk.low[node], walk.index[target])
      }

      continue
    }

    frames.pop()
    if (frames.length) {
      const parent = frames[frames.length - 1][0]
      walk.low[parent] = Math.min(walk.low[parent], walk.low[node])
    }

    if (walk.low[node] === walk.index[node]) {
      emit(walk, node)
    }
  }
}

export function order(edges: number[][]): Ordered {
  const n = edges.length
  const walk: Walk = {
    edges,
    index: new Int32Array(n).fill(-1),
    low: new Int32Array(n),
    onStack: new Uint8Array(n),
    stack: [],
    counter: 0,
    out: { order: [], cyclic: new Set() }
  }
  for (let i = 0; i < n; i++) {
    if (walk.index[i] === -1) {
      visit(walk, i)
    }
  }

  return walk.out
}
