const encoder = typeof TextEncoder === 'function' ? new TextEncoder() : null
const decoder =
  typeof TextDecoder === 'function' ? new TextDecoder('utf-8', { ignoreBOM: true }) : null

interface CodePoint {
  code: number
  width: number
}

interface Decoded {
  code: number
  next: number
}

function codePointAt(text: string, i: number): CodePoint {
  const c = text.charCodeAt(i)
  if (c < 0xd800 || c > 0xdfff) {
    return { code: c, width: 1 }
  }

  const next = text.charCodeAt(i + 1)
  if (c <= 0xdbff && next >= 0xdc00 && next <= 0xdfff) {
    return { code: 0x10000 + ((c - 0xd800) << 10) + (next - 0xdc00), width: 2 }
  }

  return { code: 0xfffd, width: 1 }
}

function writeCodePoint(out: Uint8Array, at: number, c: number): number {
  if (c < 0x80) {
    out[at] = c
    return 1
  }

  if (c < 0x800) {
    out[at] = 0xc0 | (c >> 6)
    out[at + 1] = 0x80 | (c & 0x3f)
    return 2
  }

  if (c < 0x10000) {
    out[at] = 0xe0 | (c >> 12)
    out[at + 1] = 0x80 | ((c >> 6) & 0x3f)
    out[at + 2] = 0x80 | (c & 0x3f)
    return 3
  }

  out[at] = 0xf0 | (c >> 18)
  out[at + 1] = 0x80 | ((c >> 12) & 0x3f)
  out[at + 2] = 0x80 | ((c >> 6) & 0x3f)
  out[at + 3] = 0x80 | (c & 0x3f)

  return 4
}

export function encode(text: string): Uint8Array {
  if (encoder) {
    return encoder.encode(text)
  }

  const out = new Uint8Array(text.length * 3)
  let n = 0
  for (let i = 0; i < text.length;) {
    const { code, width } = codePointAt(text, i)
    n += writeCodePoint(out, n, code)
    i += width
  }

  return out.slice(0, n)
}

function sequence(bytes: Uint8Array, at: number): Decoded {
  const b = bytes[at]
  if (b < 0x80) {
    return { code: b, next: at + 1 }
  }

  let need = 0
  let code = 0
  if (b >= 0xc2 && b < 0xe0) {
    need = 1
    code = b & 0x1f
  } else if (b >= 0xe0 && b < 0xf0) {
    need = 2
    code = b & 0x0f
  } else if (b >= 0xf0 && b < 0xf5) {
    need = 3
    code = b & 0x07
  } else {
    return { code: 0xfffd, next: at + 1 }
  }

  for (let k = 1; k <= need; k++) {
    const cont = bytes[at + k]
    if (cont === undefined || (cont & 0xc0) !== 0x80) {
      return { code: 0xfffd, next: at + k }
    }

    code = (code << 6) | (cont & 0x3f)
  }

  const overlong = (need === 2 && code < 0x800) || (need === 3 && code < 0x10000)
  const invalid = (code >= 0xd800 && code <= 0xdfff) || code > 0x10ffff

  return { code: overlong || invalid ? 0xfffd : code, next: at + need + 1 }
}

export function decode(bytes: Uint8Array): string {
  if (decoder) {
    return decoder.decode(bytes)
  }

  let out = ''
  let chunk: number[] = []
  for (let i = 0; i < bytes.length;) {
    const { code, next } = sequence(bytes, i)
    if (code >= 0x10000) {
      chunk.push(0xd800 + ((code - 0x10000) >> 10), 0xdc00 + ((code - 0x10000) & 0x3ff))
    } else {
      chunk.push(code)
    }

    if (chunk.length > 8192) {
      out += String.fromCharCode(...chunk)
      chunk = []
    }

    i = next
  }

  return out + String.fromCharCode(...chunk)
}
