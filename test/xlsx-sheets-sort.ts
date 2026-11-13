import test from 'brittle'
import * as xlsx from '../lib/xlsx/index.ts'
import { DocsError } from '../lib/errors.ts'
import { decode, unzip } from './helpers/engine.ts'

const reopen = (book: xlsx.Workbook): xlsx.Workbook => xlsx.open(book.save())
const partXml = (book: xlsx.Workbook, name: string): string => decode(unzip(book.save())[name])

function code(fn: () => unknown): string | undefined {
  try {
    fn()
  } catch (err) {
    return err instanceof DocsError ? err.code : 'not a DocsError'
  }

  return undefined
}

function people(): xlsx.Workbook {
  const book = xlsx.create()
  book.appendRows(0, [
    ['Name', 'Team', 'Score'],
    ['ana', 'red', 7],
    ['Bo', 'blue', 9],
    ['cy', 'red', null],
    ['Di', 'blue', 7],
    ['Ed', 'red', 'n/a']
  ])

  return book
}

test('autofilter sets, replaces and removes the sheet filter', (t) => {
  const book = people()
  t.alike(book.autofilter(0, 'A1:C6'), { range: 'A1:C6', changed: true })
  t.alike(book.autofilter(0, 'A1:C6'), { range: 'A1:C6', changed: false })
  const again = reopen(book)
  t.ok(partXml(again, 'xl/worksheets/sheet1.xml').includes('<autoFilter ref="A1:C6"/>'))
  t.alike(again.names(), [
    { name: '_xlnm._FilterDatabase', ref: 'Sheet1!$A$1:$C$6', sheet: 'Sheet1', hidden: true }
  ])
  t.alike(again.autofilter(0, null), { range: null, changed: true })
  t.absent(partXml(again, 'xl/worksheets/sheet1.xml').includes('autoFilter'))
  t.alike(again.names(), [])
  t.alike(again.autofilter(0, null), { range: null, changed: false })
})

test('sort orders numbers, then text, then blanks, keeping the header', (t) => {
  const book = people()
  const result = book.sort(0, {
    range: 'A1:C6',
    header: true,
    by: [{ column: 'C' }, { column: 'A' }]
  })
  t.is(result.range, 'A1:C6')
  t.is(result.rows, 5)
  t.alike(reopen(book).read(0).rows, [
    ['Name', 'Team', 'Score'],
    ['ana', 'red', 7],
    ['Di', 'blue', 7],
    ['Bo', 'blue', 9],
    ['Ed', 'red', 'n/a'],
    ['cy', 'red', null]
  ])
})

test('sort descending by several keys, case-insensitive and stable', (t) => {
  const book = people()
  book.sort(0, {
    range: 'A2:C6',
    by: [
      { column: 'B', descending: true },
      { column: 3, descending: true }
    ]
  })
  t.alike(
    reopen(book)
      .read(0, { range: 'A2:A6' })
      .rows.map((r) => r[0]),
    ['Ed', 'ana', 'cy', 'Bo', 'Di']
  )
})

test('sort moves formulas that only look outside the range', (t) => {
  const book = xlsx.create()
  book.setCells(0, { A1: 3, A2: 1, A3: 2, B1: '=$E$1*2', E1: 10 })
  book.sort(0, { range: 'A1:B3', by: [{ column: 'A' }] })
  const read = reopen(book).read(0, { range: 'A1:B3' })
  t.alike(
    read.rows.map((r) => r[0]),
    [1, 2, 3]
  )
  t.alike(read.formulas, { B3: '=$E$1*2' })
})

test('sort refuses formulas inside the range, merges and tables it straddles', (t) => {
  const book = xlsx.create()
  book.setCells(0, { A1: 3, A2: 1, B1: '=A1*2', D1: 1, D2: 2, E1: 1, E2: 2 })
  t.is(
    code(() => book.sort(0, { range: 'A1:B2', by: [{ column: 'A' }] })),
    'SORT_HAS_FORMULAS'
  )
  book.merge(0, 'D1:D2')
  t.is(
    code(() => book.sort(0, { range: 'D1:E2', by: [{ column: 'E' }] })),
    'RANGE_HAS_MERGES'
  )
  book.addTable(0, 'G1:H4')
  t.is(
    code(() => book.sort(0, { range: 'G3:I6', by: [{ column: 'G' }] })),
    'TABLE_OVERLAP'
  )
  t.is(
    code(() => book.sort(0, { range: 'G1:H4', by: [{ column: 'G' }] })),
    'TABLE_OVERLAP'
  )
  t.is(book.sort(0, { range: 'G2:H4', by: [{ column: 'G' }] }).rows, 3)
  t.is(
    code(() => book.sort(0, { range: 'A1:B2', by: [{ column: 'C' }] })),
    'INVALID_RANGE'
  )
  t.is(
    code(() => book.sort(0, { range: 'A1:B2', by: [] })),
    'INVALID_OPTIONS'
  )
  t.is(
    code(() => book.autofilter(0, 'H1:J3')),
    'TABLE_OVERLAP'
  )
})
