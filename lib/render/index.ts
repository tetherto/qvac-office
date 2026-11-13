import { DocsError } from '../errors.ts'
import { errorMessage, isRecord } from '../guards.ts'
import type { Doc } from '../tools/index.ts'

export interface RenderOptions {
  pages?: number[]
  scale?: number
}

export interface RenderedPage {
  page: number
  width: number
  height: number
  png: Uint8Array
}

export interface Rendered {
  pageCount: number
  pages: RenderedPage[]
}

type Input = Uint8Array | Doc

const DEFAULT_PAGES = 10

interface Converter {
  convert(input: Uint8Array, opts: { to: string }): Promise<Uint8Array>
}

interface Pdfium {
  open(pdf: Uint8Array): {
    pageCount(): number
    render(
      page: number,
      opts: { scale: number }
    ): { width: number; height: number; data: Uint8Array }
    close(): void
  }
}

interface Png {
  encode(image: { width: number; height: number; data: Uint8Array }): Uint8Array
}

type Guard<T> = (value: unknown) => value is T

function hasMethods(value: unknown, names: readonly string[]): boolean {
  return isRecord(value) && names.every((name) => typeof value[name] === 'function')
}

const isConverter: Guard<Converter> = (value): value is Converter => hasMethods(value, ['convert'])
const isPdfium: Guard<Pdfium> = (value): value is Pdfium => hasMethods(value, ['open'])
const isPng: Guard<Png> = (value): value is Png => hasMethods(value, ['encode'])

function unavailable(name: string, reason: string): DocsError {
  return new DocsError(
    'RENDER_UNAVAILABLE',
    `rendering needs Bare with ${name} installed (${reason})`,
    {
      missing: name
    }
  )
}

async function load<T>(name: string, guard: Guard<T>): Promise<T> {
  let mod: unknown
  try {
    mod = await import(name)
  } catch (err) {
    throw unavailable(name, errorMessage(err))
  }

  const exported = isRecord(mod) && 'default' in mod ? mod.default : mod
  if (!guard(exported)) {
    throw unavailable(name, 'it does not export the expected functions')
  }

  return exported
}

function bytesOf(input: Input): Uint8Array {
  return input instanceof Uint8Array ? input : input.save()
}

export async function toPdf(input: Input): Promise<Uint8Array> {
  const converter = await load('bare-x-to-pdf', isConverter)
  return new Uint8Array(await converter.convert(bytesOf(input), { to: 'pdf' }))
}

function checkPages(pages: number[], count: number): void {
  const bad = pages.find((page) => !Number.isInteger(page) || page < 0 || page >= count)
  if (bad !== undefined) {
    throw new DocsError('INVALID_OPTIONS', `pages must be from 0 to ${count - 1}, got ${bad}`, {
      field: 'pages',
      got: bad,
      expected: `0 to ${count - 1}`
    })
  }
}

export async function render(input: Input, opts: RenderOptions = {}): Promise<Rendered> {
  const pdf = await toPdf(input)
  const pdfium = await load('bare-pdfium', isPdfium)
  const png = await load('bare-png', isPng)
  const doc = pdfium.open(pdf)
  try {
    const pageCount = doc.pageCount()
    const wanted =
      opts.pages ?? Array.from({ length: Math.min(pageCount, DEFAULT_PAGES) }, (_, i) => i)
    checkPages(wanted, pageCount)
    const pages = wanted.map((page) => {
      const image = doc.render(page, { scale: opts.scale ?? 1 })
      return { page, width: image.width, height: image.height, png: png.encode(image) }
    })

    return { pageCount, pages }
  } finally {
    doc.close()
  }
}
