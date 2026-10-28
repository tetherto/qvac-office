import test from 'brittle'
import { pptx, docx, xlsx, Package, DocsError, CODES } from '../lib/index.ts'
import * as pptxEntry from '../lib/pptx/index.ts'
import * as docxEntry from '../lib/docx/index.ts'
import * as xlsxEntry from '../lib/xlsx/index.ts'
import { fixture } from './helpers/engine.ts'

function codeOf(fn: () => unknown): string | null {
  try {
    fn()
  } catch (err) {
    return err instanceof DocsError ? err.code : 'NOT_DOCS_ERROR'
  }

  return null
}

test('each format has its own entry point, and the engine entry re-exports them', (t) => {
  t.is(pptxEntry.open, pptx.open)
  t.is(docxEntry.create, docx.create)
  t.is(xlsxEntry.Workbook, xlsx.Workbook)
})

test('every error code is registered and described', (t) => {
  t.ok(Object.keys(CODES).length > 40)
  for (const [code, description] of Object.entries(CODES)) {
    t.ok(description.length > 10, code)
  }

  t.is(DocsError.codes, CODES)
})

test('strict mode catches a part changed without touch()', (t) => {
  const deck = pptx.open(fixture('basic.pptx'))
  deck.package.root('/ppt/slides/slide1.xml').setAttribute('show', '0')
  t.is(
    codeOf(() => deck.save()),
    'UNTOUCHED_EDIT'
  )
  deck.package.touch('/ppt/slides/slide1.xml')
  t.execution(() => deck.save())
})

test('strict mode is off unless asked for', (t) => {
  const pkg = Package.open(fixture('basic.pptx'), { strict: false })
  pkg.root('/ppt/presentation.xml').setAttribute('x', '1')
  t.execution(() => pkg.save())
})

test('format is the shared name for direct formatting', (t) => {
  const deck = pptx.open(fixture('basic.pptx'))
  t.alike(deck.format({ slide: 0, shape: 'title' }, { bold: true }), { runs: 1 })
  const doc = docx.open(fixture('report.docx'))
  t.alike(doc.format(1, { italic: true }), { changed: true })
  const book = xlsx.open(fixture('shared.xlsx'))
  t.alike(book.format(0, 'A2:B3', { fill: 'FFF2CC' }), { cells: 4 })
  const again = xlsx.open(book.save())
  t.is(again.get(0, 'B3').value, 450.5)
  t.is(again.get(0, 'B3').format, '#,##0.00')
  t.exception(() => book.format(0, 'A1', {}), /must name something to change/)
  const loose: Record<string, unknown> = { value: 1 }
  t.exception(() => book.format(0, 'A1', loose), /unknown cell style "value"/)
})

test('read pages through large documents and decks', (t) => {
  const doc = docx.open(fixture('report.docx'))
  const page = doc.read({ from: 4, count: 3 })
  t.is(page.total, 14)
  t.alike(
    page.blocks.map((b) => b.index),
    [4, 5, 6]
  )
  t.is(doc.read({ from: 12 }).blocks.length, 2)
  t.exception(() => doc.read({ from: 14 }), /from must be a block index from 0 to 13/)
  t.exception(() => doc.read({ count: 0 }), /count must be at least 1/)
  const deck = pptx.open(fixture('basic.pptx'))
  t.alike(
    deck.read({ from: 1, count: 1 }).map((s) => s.index),
    [1]
  )
  t.alike(
    deck.read({ slides: [2, 0] }).map((s) => s.index),
    [0, 2]
  )
  t.exception(() => deck.read({ slides: [5] }), /slides must list slide indexes from 0 to 2/)
  t.exception(() => deck.read({ slides: [0], from: 1 }), /either slides or from\/count/)
  t.is(pptx.create().read().length, 0)
})
