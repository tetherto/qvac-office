import { probe, type ImageInfo } from '../image.ts'
import * as paths from './paths.ts'
import type { Package } from './package.ts'

export const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

export interface ImagePlacement {
  source: string
  dir: string
}

export interface AddedImage {
  rId: string
  part: string
  info: ImageInfo
}

export function addImage(pkg: Package, bytes: Uint8Array, at: ImagePlacement): AddedImage {
  const info = probe(bytes)
  pkg.contentTypes.addDefault(info.format, info.contentType)
  const part = pkg.freeName(paths.resolve(at.source, `${at.dir}/image%d.${info.format}`))
  pkg.add(part, bytes)
  const rId = pkg.rels(at.source).add(IMAGE_REL, part)
  return { rId, part, info }
}
