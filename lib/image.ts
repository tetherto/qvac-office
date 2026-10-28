import { DocsError } from './errors.ts'

export type ImageFormat = 'png' | 'jpeg' | 'gif'

export interface ImageInfo {
  format: ImageFormat
  contentType: string
  width: number
  height: number
}

function jpegSize(bytes: Uint8Array, view: DataView): ImageInfo {
  let offset = 2
  while (offset + 9 < bytes.byteLength && bytes[offset] === 0xff) {
    const marker = bytes[offset + 1]
    if (marker === 0xff) {
      offset++
      continue
    }

    const frame = marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)
    if (frame) {
      return {
        format: 'jpeg',
        contentType: 'image/jpeg',
        width: view.getUint16(offset + 7),
        height: view.getUint16(offset + 5)
      }
    }

    offset += 2 + view.getUint16(offset + 2)
  }

  throw new DocsError('INVALID_IMAGE', 'the JPEG has no frame header, so its size is unknown')
}

export function probe(bytes: Uint8Array): ImageInfo {
  if (!(bytes instanceof Uint8Array) || bytes.byteLength < 24) {
    throw new DocsError('INVALID_IMAGE', 'the image is empty or truncated')
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (view.getUint32(0) === 0x89504e47 && view.getUint32(4) === 0x0d0a1a0a) {
    return {
      format: 'png',
      contentType: 'image/png',
      width: view.getUint32(16),
      height: view.getUint32(20)
    }
  }

  if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) {
    return {
      format: 'gif',
      contentType: 'image/gif',
      width: view.getUint16(6, true),
      height: view.getUint16(8, true)
    }
  }

  if (bytes[0] === 0xff && bytes[1] === 0xd8) {
    return jpegSize(bytes, view)
  }

  throw new DocsError('UNSUPPORTED_IMAGE', 'only PNG, JPEG and GIF images can be embedded')
}
