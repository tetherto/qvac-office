import { DocsError } from '../errors.ts'

const ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'
const LOOKUP = new Map([...ALPHABET].map((char, i) => [char, i]))
const DATA_URL = /^data:[^,]*;base64,/

export function decode(input: string, field = 'image'): Uint8Array {
  const text = input
    .replace(DATA_URL, '')
    .replace(/[\s=]/g, '')
    .replace(/-/g, '+')
    .replace(/_/g, '/')
  const bytes = new Uint8Array(Math.floor((text.length * 3) / 4))
  let length = 0
  let buffer = 0
  let bits = 0
  for (const char of text) {
    const value = LOOKUP.get(char)
    if (value === undefined) {
      throw new DocsError('INVALID_IMAGE', `${field} must be base64 or a base64 data: URL`, {
        field,
        expected: 'base64'
      })
    }

    buffer = (buffer << 6) | value
    bits += 6
    if (bits >= 8) {
      bits -= 8
      bytes[length++] = (buffer >> bits) & 0xff
    }
  }

  return bytes.subarray(0, length)
}
