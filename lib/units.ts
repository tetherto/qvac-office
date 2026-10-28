const EMU_PER_INCH = 914400
const EMU_PER_POINT = 12700
const EMU_PER_CM = 360000

export function inches(n: number): number {
  return Math.round(n * EMU_PER_INCH)
}

export function points(n: number): number {
  return Math.round(n * EMU_PER_POINT)
}

export function cm(n: number): number {
  return Math.round(n * EMU_PER_CM)
}
