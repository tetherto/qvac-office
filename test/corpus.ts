import test from 'brittle'
import b4a from 'b4a'
import { Package, DocsError } from '../lib/index.ts'
import { run, expected, readCorpus } from './helpers/corpus.ts'
import { unzip } from './helpers/engine.ts'

test('every corpus file opens, edits and saves, or fails with its recorded DocsError', (t) => {
  t.alike(run(), expected())
})

test('an untouched corpus file saves every part byte for byte', (t) => {
  for (const [name, outcome] of Object.entries(expected())) {
    if (outcome !== 'ok') continue
    const input = readCorpus(name)
    let saved: Uint8Array
    try {
      saved = Package.open(input).save()
    } catch (err) {
      t.fail(`${name}: ${err instanceof DocsError ? err.code : String(err)}`)
      continue
    }

    const before = unzip(input)
    const after = unzip(saved)
    const changed = Object.keys(before).filter(
      (part) => !after[part] || !b4a.equals(after[part], before[part])
    )
    t.alike(changed, [], name)
  }
})
