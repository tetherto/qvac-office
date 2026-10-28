import fs from 'fs'
import process from 'process'
import { run, expected, CORPUS } from '../../test/helpers/corpus.ts'

const results = run()
if (process.argv.includes('--write')) {
  fs.writeFileSync(new URL('expected.json', CORPUS), JSON.stringify(results, null, 2) + '\n')
} else {
  const recorded = expected()
  for (const [name, outcome] of Object.entries(results)) {
    if (recorded[name] !== outcome) {
      console.log(`${name}: expected ${recorded[name]}, got ${outcome}`)
    }
  }
}
