import test from 'brittle'
import * as xlsx from '../lib/xlsx/index.ts'
import { DocsError } from '../lib/errors.ts'
import { decode, unzip } from './helpers/engine.ts'

const reopen = (book: xlsx.Workbook): xlsx.Workbook => xlsx.open(book.save())
const sheetXml = (book: xlsx.Workbook): string =>
  decode(unzip(book.save())['xl/worksheets/sheet1.xml'])

function code(fn: () => unknown): string | undefined {
  try {
    fn()
  } catch (err) {
    return err instanceof DocsError ? err.code : 'not a DocsError'
  }

  return undefined
}

test('merge keeps the top-left value and reports the cells it cleared', (t) => {
  const book = xlsx.create()
  book.setCells(0, { A1: 'Title', B1: 'gone', A2: 3, C2: 4, D1: 'kept' })
  t.alike(book.merge(0, 'A1:C2'), { range: 'A1:C2', cleared: ['B1', 'A2', 'C2'] })
  const again = reopen(book)
  const read = again.read(0)
  t.alike(read.merged, ['A1:C2'])
  t.alike(read.rows[0], ['Title', null, null, 'kept'])
  t.is(
    code(() => again.merge(0, 'C2:D3')),
    'MERGE_OVERLAP'
  )
  t.is(
    code(() => again.merge(0, 'E5')),
    'INVALID_RANGE'
  )
  t.is(
    code(() => again.setCells(0, { B2: 1 })),
    'CELL_MERGED'
  )
})

test('unmerge drops every merge the range touches', (t) => {
  const book = xlsx.create()
  book.merge(0, 'A1:B1')
  book.merge(0, 'D1:E2')
  book.merge(0, 'A5:A6')
  t.alike(book.unmerge(0, 'B1:D1'), { unmerged: ['A1:B1', 'D1:E2'] })
  t.alike(book.unmerge(0, 'H9'), { unmerged: [] })
  t.alike(reopen(book).read(0).merged, ['A5:A6'])
  book.unmerge(0, 'A5')
  t.absent(sheetXml(book).includes('mergeCells'))
})

test('setRowHeight takes points or a length', (t) => {
  const book = xlsx.create()
  book.setCells(0, { A2: 'x' })
  t.alike(book.setRowHeight(0, { row: 2, height: 30 }), { row: 2, height: 30 })
  t.alike(book.setRowHeight(0, { row: 5, height: '0.5in' }), { row: 5, height: 36 })
  const sheet = sheetXml(book)
  t.ok(sheet.includes('<row r="2" ht="30" customHeight="1">'))
  t.ok(sheet.includes('<row r="5" ht="36" customHeight="1"/>'))
  t.is(
    code(() => book.setRowHeight(0, { row: 1, height: 500 })),
    'INVALID_VALUE'
  )
  t.is(
    code(() => book.setRowHeight(0, { row: 0, height: 10 })),
    'INVALID_CELL'
  )
})

test('hide and unhide rows and columns', (t) => {
  const book = xlsx.create()
  book.setCells(0, { A1: 1, A3: 3 })
  book.setColumnWidth(0, { column: 'C', width: 12 })
  t.alike(book.hide(0, { rows: [2, 3], columns: ['B', 'D'] }), {
    rows: 2,
    columns: 3,
    hidden: true
  })
  const hidden = sheetXml(reopen(book))
  t.ok(hidden.includes('<row r="2" hidden="1"/>'))
  t.ok(hidden.includes('<row r="3" hidden="1">'))
  t.ok(hidden.includes('<col min="2" max="2" width="9.140625" hidden="1"/>'))
  t.ok(hidden.includes('<col min="3" max="3" width="12" customWidth="1" hidden="1"/>'))
  t.ok(hidden.includes('<col min="4" max="4" width="9.140625" hidden="1"/>'))
  book.hide(0, { rows: 3, columns: 'C', hidden: false })
  const shown = sheetXml(book)
  t.ok(shown.includes('<row r="3">'))
  t.ok(shown.includes('<col min="3" max="3" width="12" customWidth="1"/>'))
  t.ok(shown.includes('<col min="2" max="2" width="9.140625" hidden="1"/>'))
  t.is(
    code(() => book.hide(0, {})),
    'INVALID_OPTIONS'
  )
  t.is(
    code(() => book.hide(0, { rows: [5, 2] })),
    'INVALID_RANGE'
  )
})

test('freeze writes a frozen pane and 0/0 removes it', (t) => {
  const book = xlsx.create()
  t.alike(book.freeze(0, { rows: 1 }), { rows: 1, columns: 0, topLeftCell: 'A2' })
  t.ok(
    sheetXml(book).includes(
      '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
    )
  )
  book.freeze(0, { rows: 2, columns: 1 })
  const both = sheetXml(reopen(book))
  t.ok(both.includes('xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight"'))
  t.is(both.split('<pane').length, 2)
  t.alike(book.freeze(0, { rows: 0, columns: 0 }), { rows: 0, columns: 0, topLeftCell: null })
  t.absent(sheetXml(book).includes('<pane'))
  t.is(
    code(() => book.freeze(0, { rows: -1 })),
    'INVALID_VALUE'
  )
})
