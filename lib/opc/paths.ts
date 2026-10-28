export function dirname(partName: string): string {
  return partName.slice(0, partName.lastIndexOf('/') + 1)
}

export function basename(partName: string): string {
  return partName.slice(partName.lastIndexOf('/') + 1)
}

export function extension(partName: string): string {
  const base = basename(partName)
  const dot = base.lastIndexOf('.')
  return dot === -1 ? '' : base.slice(dot + 1).toLowerCase()
}

export function normalize(path: string): string {
  const out: string[] = []
  for (const segment of path.split('/')) {
    if (segment === '' || segment === '.') continue
    if (segment === '..') {
      out.pop()
    } else {
      out.push(segment)
    }
  }

  return '/' + out.join('/')
}

export function relsFor(source: string): string {
  if (source === '/') return '/_rels/.rels'
  return `${dirname(source)}_rels/${basename(source)}.rels`
}

export function resolve(source: string, target: string): string {
  const clean = target.split('#')[0]
  if (clean.startsWith('/')) {
    return normalize(clean)
  }

  return normalize(dirname(source) + clean)
}

export function relative(source: string, target: string): string {
  const from = dirname(source).split('/').filter(Boolean)
  const to = target.split('/').filter(Boolean)
  let common = 0
  while (common < from.length && common < to.length - 1 && from[common] === to[common]) {
    common++
  }

  return [...from.slice(common).map(() => '..'), ...to.slice(common)].join('/')
}
