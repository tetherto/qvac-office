import * as xml from '../xml.ts'
import type { Document } from '../xml.ts'

const { p: P } = xml.ns

export function dropAnimations(doc: Document, ids: Set<string>): boolean {
  const timing = xml.child(xml.root(doc), P, 'timing')
  if (!timing || !xml.descendants(timing, P, 'spTgt').some((t) => ids.has(xml.attr(t, 'spid')))) {
    return false
  }

  xml.remove(timing)

  return true
}
