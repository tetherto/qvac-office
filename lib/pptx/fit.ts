import type { Box, Insets } from './inherit.ts'

export interface SizedParagraph {
  text: string
  level: number
  size: number
}

export interface Estimate {
  overflow: boolean
  fontScale: number
  lineReduction: number
}

interface Squeeze {
  scale: number
  reduction: number
}

const EMU_PER_POINT = 12700
const LEVEL_INDENT = 342900
const FLOOR: Estimate = { overflow: true, fontScale: 0.25, lineReduction: 0.2 }

function paragraphHeight(p: SizedParagraph, width: number, squeeze: Squeeze): number {
  const size = p.size * squeeze.scale * EMU_PER_POINT
  const available = Math.max(width - p.level * LEVEL_INDENT, width * 0.3)
  const perLine = Math.max(1, Math.floor(available / (size * 0.5)))
  let lines = 0
  for (const line of p.text.split('\n')) {
    lines += Math.max(1, Math.ceil(line.length / perLine))
  }

  return lines * size * 1.2 * (1 - squeeze.reduction)
}

function neededHeight(paragraphs: SizedParagraph[], width: number, squeeze: Squeeze): number {
  return paragraphs.reduce((total, p) => total + paragraphHeight(p, width, squeeze), 0)
}

export function estimate(paragraphs: SizedParagraph[], box: Box, insets: Insets): Estimate {
  const width = box.width - insets.left - insets.right
  const height = box.height - insets.top - insets.bottom
  if (width <= 0 || height <= 0) {
    return FLOOR
  }

  if (neededHeight(paragraphs, width, { scale: 1, reduction: 0 }) <= height) {
    return { overflow: false, fontScale: 1, lineReduction: 0 }
  }

  for (let scale = 0.975; scale >= 0.25; scale -= 0.025) {
    const reduction = scale < 0.7 ? 0.2 : 0.1
    if (neededHeight(paragraphs, width, { scale, reduction }) <= height) {
      return {
        overflow: true,
        fontScale: Math.round(scale * 1000) / 1000,
        lineReduction: reduction
      }
    }
  }

  return FLOOR
}
