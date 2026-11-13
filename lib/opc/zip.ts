import { deflateRawSync, inflateRawSync } from '#zlib'
import { DocsError } from '../errors.ts'
import { errorCode, errorMessage } from '../guards.ts'
import * as utf8 from '../utf8.ts'

export interface Limits {
  maxEntries: number
  maxSize: number
  maxRatio: number
}

export interface ZipEntry {
  name: string
  data: Uint8Array
}

export interface ZipInput extends ZipEntry {
  compress: boolean
}

interface Directory {
  offset: number
  count: number
}

interface CentralEntry {
  name: string
  method: number
  flags: number
  crc: number
  compressed: number
  size: number
  offset: number
  next: number
}

const EOCD = 0x06054b50
const EOCD64_LOCATOR = 0x07064b50
const EOCD64 = 0x06064b50
const CENTRAL = 0x02014b50
const LOCAL = 0x04034b50
const DOS_DATE = 33
const UTF8_FLAG = 0x0800

const CRC_TABLE = new Uint32Array(256).map((_, n) => {
  let c = n
  for (let k = 0; k < 8; k++) {
    c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
  }

  return c >>> 0
})

export function crc32(data: Uint8Array): number {
  let crc = 0xffffffff
  for (let i = 0; i < data.length; i++) {
    crc = CRC_TABLE[(crc ^ data[i]) & 0xff] ^ (crc >>> 8)
  }

  return (crc ^ 0xffffffff) >>> 0
}

function broken(why: string): DocsError {
  return new DocsError('NOT_OOXML', `the file is not a readable zip archive: ${why}`)
}

function findDirectory(view: DataView): Directory {
  const min = Math.max(0, view.byteLength - 65557)
  for (let at = view.byteLength - 22; at >= min; at--) {
    if (view.getUint32(at, true) !== EOCD) continue
    const count = view.getUint16(at + 10, true)
    const offset = view.getUint32(at + 16, true)
    const locator = at - 20
    if (
      (count === 0xffff || offset === 0xffffffff) &&
      locator >= 0 &&
      view.getUint32(locator, true) === EOCD64_LOCATOR
    ) {
      return zip64Directory(view, Number(view.getBigUint64(locator + 8, true)))
    }

    return { offset, count }
  }

  throw broken('no end of central directory')
}

function zip64Directory(view: DataView, at: number): Directory {
  if (at + 56 > view.byteLength || view.getUint32(at, true) !== EOCD64) {
    throw broken('bad zip64 directory')
  }

  return {
    count: Number(view.getBigUint64(at + 32, true)),
    offset: Number(view.getBigUint64(at + 48, true))
  }
}

interface Span {
  start: number
  end: number
}

function zip64Extra(view: DataView, entry: CentralEntry, extra: Span): void {
  for (let at = extra.start; at + 4 <= extra.end;) {
    const id = view.getUint16(at, true)
    const length = view.getUint16(at + 2, true)
    if (id === 0x0001) {
      let field = at + 4
      const take = (): number => {
        if (field + 8 > extra.end) {
          throw broken(`${entry.name} has a short zip64 field`)
        }

        const value = Number(view.getBigUint64(field, true))
        field += 8

        return value
      }

      if (entry.size === 0xffffffff) {
        entry.size = take()
      }

      if (entry.compressed === 0xffffffff) {
        entry.compressed = take()
      }

      if (entry.offset === 0xffffffff) {
        entry.offset = take()
      }

      return
    }

    at += 4 + length
  }
}

function readCentral(bytes: Uint8Array, view: DataView, at: number): CentralEntry {
  if (at + 46 > view.byteLength || view.getUint32(at, true) !== CENTRAL) {
    throw broken('bad central directory entry')
  }

  const nameLength = view.getUint16(at + 28, true)
  const extraLength = view.getUint16(at + 30, true)
  const commentLength = view.getUint16(at + 32, true)
  const entry: CentralEntry = {
    name: utf8.decode(bytes.subarray(at + 46, at + 46 + nameLength)),
    flags: view.getUint16(at + 8, true),
    method: view.getUint16(at + 10, true),
    crc: view.getUint32(at + 16, true),
    compressed: view.getUint32(at + 20, true),
    size: view.getUint32(at + 24, true),
    offset: view.getUint32(at + 42, true),
    next: at + 46 + nameLength + extraLength + commentLength
  }
  const start = at + 46 + nameLength
  zip64Extra(view, entry, { start, end: start + extraLength })

  return entry
}

function checkLimits(entries: CentralEntry[], limits: Limits): void {
  if (entries.length > limits.maxEntries) {
    throw new DocsError('TOO_LARGE', `the file has more than ${limits.maxEntries} entries`)
  }

  let declared = 0
  for (const entry of entries) {
    declared += entry.size
    if (declared > limits.maxSize) {
      throw new DocsError('TOO_LARGE', `the file expands past ${limits.maxSize} bytes`)
    }

    if (entry.size > 1024 * 1024 && entry.size / Math.max(1, entry.compressed) > limits.maxRatio) {
      throw new DocsError(
        'TOO_LARGE',
        `${entry.name} expands more than ${limits.maxRatio} times, which looks like a zip bomb`
      )
    }
  }
}

function extract(bytes: Uint8Array, view: DataView, entry: CentralEntry): Uint8Array {
  const at = entry.offset
  if (at + 30 > view.byteLength || view.getUint32(at, true) !== LOCAL) {
    throw broken(`bad local header for ${entry.name}`)
  }

  if (entry.flags & 1) {
    throw broken(`${entry.name} is encrypted`)
  }

  const start = at + 30 + view.getUint16(at + 26, true) + view.getUint16(at + 28, true)
  const raw = bytes.subarray(start, start + entry.compressed)
  if (raw.length !== entry.compressed) {
    throw broken(`${entry.name} is truncated`)
  }

  let data: Uint8Array
  if (entry.method === 0) {
    data = raw
  } else if (entry.method === 8) {
    data = inflate(raw, entry)
  } else {
    throw broken(`${entry.name} uses compression method ${entry.method}`)
  }

  if (data.length !== entry.size) {
    throw broken(`${entry.name} has the wrong size`)
  }

  if (crc32(data) !== entry.crc) {
    throw broken(`${entry.name} fails its checksum`)
  }

  return data
}

function inflate(raw: Uint8Array, entry: CentralEntry): Uint8Array {
  try {
    return new Uint8Array(inflateRawSync(raw, { maxOutputLength: Math.max(1, entry.size) }))
  } catch (err) {
    if (isOverflow(err)) {
      throw broken(`${entry.name} inflates past its declared size`)
    }

    throw broken(`${entry.name} does not inflate`)
  }
}

function isOverflow(err: unknown): boolean {
  const code = errorCode(err)
  return code === 'ERR_BUFFER_TOO_LARGE' || code === 'LIMIT_EXCEEDED'
}

export function read(bytes: Uint8Array, limits: Limits): ZipEntry[] {
  try {
    return readEntries(bytes, limits)
  } catch (err) {
    if (err instanceof DocsError) {
      throw err
    }

    throw broken(errorMessage(err))
  }
}

function readEntries(bytes: Uint8Array, limits: Limits): ZipEntry[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (bytes.byteLength < 22) {
    throw broken('too short')
  }

  const directory = findDirectory(view)
  const entries: CentralEntry[] = []
  for (let i = 0, at = directory.offset; i < directory.count; i++) {
    const entry = readCentral(bytes, view, at)
    entries.push(entry)
    at = entry.next
  }

  checkLimits(entries, limits)

  return entries
    .filter((entry) => !entry.name.endsWith('/'))
    .map((entry) => ({ name: entry.name, data: extract(bytes, view, entry) }))
}

interface Encoded {
  name: Uint8Array
  data: Uint8Array
  raw: Uint8Array
  crc: number
  method: number
}

function encodeEntry(entry: ZipInput): Encoded {
  const compressed = entry.compress && entry.data.length > 0
  return {
    name: utf8.encode(entry.name),
    data: entry.data,
    raw: compressed ? new Uint8Array(deflateRawSync(entry.data, { level: 6 })) : entry.data,
    crc: crc32(entry.data),
    method: compressed ? 8 : 0
  }
}

function header(size: number, fill: (view: DataView) => void): Uint8Array {
  const out = new Uint8Array(size)
  fill(new DataView(out.buffer))
  return out
}

function localHeader(e: Encoded): Uint8Array {
  return header(30, (v) => {
    v.setUint32(0, LOCAL, true)
    v.setUint16(4, 20, true)
    v.setUint16(6, UTF8_FLAG, true)
    v.setUint16(8, e.method, true)
    v.setUint16(12, DOS_DATE, true)
    v.setUint32(14, e.crc, true)
    v.setUint32(18, e.raw.length, true)
    v.setUint32(22, e.data.length, true)
    v.setUint16(26, e.name.length, true)
  })
}

function centralHeader(e: Encoded, offset: number): Uint8Array {
  return header(46, (v) => {
    v.setUint32(0, CENTRAL, true)
    v.setUint16(4, 20, true)
    v.setUint16(6, 20, true)
    v.setUint16(8, UTF8_FLAG, true)
    v.setUint16(10, e.method, true)
    v.setUint16(14, DOS_DATE, true)
    v.setUint32(16, e.crc, true)
    v.setUint32(20, e.raw.length, true)
    v.setUint32(24, e.data.length, true)
    v.setUint16(28, e.name.length, true)
    v.setUint32(42, offset, true)
  })
}

function concat(parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let at = 0
  for (const part of parts) {
    out.set(part, at)
    at += part.length
  }

  return out
}

export function write(entries: ZipInput[]): Uint8Array {
  const encoded = entries.map(encodeEntry)
  const body: Uint8Array[] = []
  const central: Uint8Array[] = []
  let offset = 0
  for (const e of encoded) {
    central.push(centralHeader(e, offset), e.name)
    body.push(localHeader(e), e.name, e.raw)
    offset += 30 + e.name.length + e.raw.length
  }

  const directory = concat(central)
  const end = header(22, (v) => {
    v.setUint32(0, EOCD, true)
    v.setUint16(8, encoded.length, true)
    v.setUint16(10, encoded.length, true)
    v.setUint32(12, directory.length, true)
    v.setUint32(16, offset, true)
  })

  return concat([...body, directory, end])
}
