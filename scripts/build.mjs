import { execFileSync } from 'child_process'
import fs from 'fs'
import { createRequire } from 'module'
import path from 'path'
import { fileURLToPath } from 'url'

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const dist = path.join(root, 'dist')

fs.rmSync(dist, { recursive: true, force: true })
const tsc = path.join(
  path.dirname(createRequire(import.meta.url).resolve('typescript/package.json')),
  'bin',
  'tsc'
)
execFileSync(process.execPath, [tsc, '-p', 'tsconfig.build.json'], { cwd: root, stdio: 'inherit' })

const count = (ext) =>
  fs.readdirSync(dist, { recursive: true }).filter((f) => String(f).endsWith(ext)).length
console.log(`[build] ${count('.js')} modules, ${count('.d.ts')} declarations -> dist/`)
