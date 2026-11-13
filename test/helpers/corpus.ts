import fs from 'fs'
import { pptx, docx, xlsx, DocsError } from '../../lib/index.ts'
import { errorMessage } from '../../lib/guards.ts'

export type Format = 'pptx' | 'docx' | 'xlsx'

export const CORPUS = new URL('../fixtures/corpus/', import.meta.url)
export const FORMATS: readonly Format[] = ['pptx', 'docx', 'xlsx']

function tolerate(fn: () => unknown, code: string): void {
  try {
    fn()
  } catch (err) {
    if (!(err instanceof DocsError) || err.code !== code) {
      throw err
    }
  }
}

function editDeck(bytes: Uint8Array): Uint8Array {
  const deck = pptx.open(bytes)
  for (const slide of deck.read()) {
    const shape = slide.shapes.find((s) => s.kind === 'text' && s.id !== null)
    if (shape?.id !== undefined && shape.id !== null) {
      deck.setText({ slide: slide.index, shape: shape.id }, 'Edited by the engine')
      break
    }
  }

  deck.addSlide({ title: 'Added by the engine' })
  const out = deck.save()
  pptx.open(out).read()

  return out
}

function editDocument(bytes: Uint8Array): Uint8Array {
  const doc = docx.open(bytes)
  const plain = doc.read().blocks.find((b) => b.type === 'paragraph' && b.text && !b.images)
  if (plain) {
    tolerate(() => doc.setText(plain.index, 'Edited by the engine'), 'TRACKED_CHANGES')
  }

  doc.appendParagraphs('Added by the engine')
  const out = doc.save()
  docx.open(out).read()

  return out
}

function editWorkbook(bytes: Uint8Array): Uint8Array {
  const book = xlsx.open(bytes)
  const sheet = book.sheets().find((s) => s.kind === 'worksheet')
  if (sheet) {
    book.read(sheet.index)
    book.appendRows(sheet.index, [['Added by the engine', 1]])
    tolerate(() => {
      book.insertRows(sheet.index, { at: 2, rows: [['Inserted by the engine', 2]] })
      book.deleteRows(sheet.index, { at: 3 })
    }, 'SHEET_HAS_DEPENDENTS')
  }

  book.addSheet('Engine')
  const out = book.save()
  xlsx.open(out).sheets()

  return out
}

export const edits: Record<Format, (bytes: Uint8Array) => Uint8Array> = {
  pptx: editDeck,
  docx: editDocument,
  xlsx: editWorkbook
}

export function corpusFiles(format: Format): string[] {
  return fs.readdirSync(new URL(`${format}/`, CORPUS)).sort()
}

export function readCorpus(name: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(new URL(name, CORPUS)))
}

export function outcome(format: Format, bytes: Uint8Array): string {
  try {
    edits[format](bytes)
    return 'ok'
  } catch (err) {
    return err instanceof DocsError ? err.code : `CRASH ${errorMessage(err)}`
  }
}

export function run(): Record<string, string> {
  const results: Record<string, string> = {}
  for (const format of FORMATS) {
    for (const name of corpusFiles(format)) {
      results[`${format}/${name}`] = outcome(format, readCorpus(`${format}/${name}`))
    }
  }

  return results
}

export function expected(): Record<string, string> {
  const parsed: unknown = JSON.parse(fs.readFileSync(new URL('expected.json', CORPUS), 'utf8'))
  const out: Record<string, string> = {}
  if (typeof parsed !== 'object' || parsed === null) {
    return out
  }

  for (const [name, value] of Object.entries(parsed)) {
    out[name] = String(value)
  }

  return out
}
