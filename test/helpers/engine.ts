import fs from 'fs'
import { Package } from '../../lib/opc/package.ts'
import * as zip from '../../lib/opc/zip.ts'
import * as utf8 from '../../lib/utf8.ts'

Package.strict = true

const LIMITS = { maxEntries: 100000, maxSize: 2 ** 31, maxRatio: 1e9 }

export function fixture(name: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(new URL(`../fixtures/engine/${name}`, import.meta.url)))
}

export function corpusFile(path: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(new URL(`../fixtures/corpus/${path}`, import.meta.url)))
}

export function decode(bytes: Uint8Array | undefined): string {
  return bytes ? utf8.decode(bytes) : ''
}

export function unzip(bytes: Uint8Array): Record<string, Uint8Array> {
  return Object.fromEntries(zip.read(bytes, LIMITS).map((e) => [e.name, e.data]))
}

export function zipFiles(files: Record<string, Uint8Array>): Uint8Array {
  return zip.write(Object.entries(files).map(([name, data]) => ({ name, data, compress: true })))
}
