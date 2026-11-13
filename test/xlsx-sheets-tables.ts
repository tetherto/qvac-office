import test from 'brittle'
import * as xlsx from '../lib/xlsx/index.ts'
import { DocsError } from '../lib/errors.ts'
import { corpusFile, decode, unzip } from './helpers/engine.ts'

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

test('defined names are listed, replaced and deleted per scope', (t) => {
  const book = xlsx.create({ sheets: ['Sales', 'Other sheet'] })
  t.alike(book.defineName('TaxRate', '=0.2'), { name: 'TaxRate', ref: '0.2', replaced: false })
  t.alike(book.defineName('Region', 'A1:B5', { sheet: 'Other sheet' }), {
    name: 'Region',
    ref: "'Other sheet'!$A$1:$B$5",
    sheet: 'Other sheet',
    replaced: false
  })
  book.defineName('Region', 'Sales!$C$1')
  t.alike(book.defineName('taxrate', '0.25').replaced, true)
  const again = reopen(book)
  t.alike(again.names(), [
    { name: 'TaxRate', ref: '0.25' },
    { name: 'Region', ref: "'Other sheet'!$A$1:$B$5", sheet: 'Other sheet' },
    { name: 'Region', ref: 'Sales!$C$1' }
  ])
  t.alike(again.deleteName('Region', { sheet: 1 }), { deleted: 1 })
  t.is(
    code(() => again.deleteName('Region', { sheet: 1 })),
    'NAME_NOT_FOUND'
  )
  again.deleteName('Region')
  again.deleteName('TAXRATE')
  t.absent(partXml(again, 'xl/workbook.xml').includes('definedNames'))
})

test('defined names follow Excel naming rules', (t) => {
  const book = xlsx.create()
  for (const bad of [
    'A1',
    'xfd100',
    'R1C1',
    'r',
    'C',
    '1st',
    'has space',
    '_xlnm.Print_Area',
    '',
    'a-b'
  ]) {
    t.is(
      code(() => book.defineName(bad, 'Sheet1!$A$1')),
      'INVALID_DEFINED_NAME',
      bad
    )
  }

  for (const good of ['_total', '\\path', 'Sales.2026', 'Ünits', 'ABCD1', 'R1C1x']) {
    t.is(
      code(() => book.defineName(good, 'Sheet1!$A$1')),
      undefined,
      good
    )
  }

  t.is(
    code(() => book.defineName('Loose', 'A1')),
    'INVALID_RANGE'
  )
  t.is(
    code(() => book.defineName('Gone', 'Missing!A1')),
    'SHEET_NOT_FOUND'
  )
  t.is(
    code(() => book.defineName('Scoped', 'A1', { sheet: 'Nope' })),
    'SHEET_NOT_FOUND'
  )
})

test('addTable builds a table part from the header row', (t) => {
  const book = xlsx.create()
  book.appendRows(0, [
    ['Item', 2026, null, 'Item'],
    ['Rent', 100, 1, 2],
    ['Food', 50, 3, 4]
  ])
  t.alike(book.addTable(0, 'A1:D3', { totals: true }), {
    name: 'Table1',
    range: 'A1:D4',
    columns: ['Item', '2026', 'Column3', 'Item2']
  })
  const again = reopen(book)
  const table = partXml(again, 'xl/tables/table1.xml')
  t.ok(
    table.includes(
      '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:D4" totalsRowCount="1"><autoFilter ref="A1:D3"/>'
    )
  )
  t.ok(table.includes('<tableColumn id="1" name="Item" totalsRowLabel="Total"/>'))
  t.ok(table.includes('<tableColumn id="4" name="Item2" totalsRowFunction="sum"/>'))
  t.ok(table.includes('<tableStyleInfo name="TableStyleMedium2"'))
  t.ok(partXml(again, '[Content_Types].xml').includes('/xl/tables/table1.xml'))
  const read = again.read(0)
  t.alike(read.rows[0], ['Item', '2026', 'Column3', 'Item2'])
  t.alike(read.rows[3], ['Total', null, null, null])
  t.alike(read.formulas, { D4: '=SUBTOTAL(109,Table1[Item2])' })
  t.alike(again.tables(), [
    {
      name: 'Table1',
      sheet: 'Sheet1',
      range: 'A1:D4',
      header: true,
      totals: true,
      columns: ['Item', '2026', 'Column3', 'Item2'],
      style: 'TableStyleMedium2'
    }
  ])
  t.is(
    again.addTable(0, 'F1:G2', { name: 'Costs', header: false, style: 'TableStyleLight9' }).name,
    'Costs'
  )
  t.ok(
    partXml(again, 'xl/tables/table2.xml').includes(
      'id="2" name="Costs" displayName="Costs" ref="F1:G2" headerRowCount="0"'
    )
  )
  t.is(again.tables(0).length, 2)
})

test('addTable refuses overlaps and taken names', (t) => {
  const book = xlsx.create()
  book.addTable(0, 'A1:B3')
  book.merge(0, 'D1:E1')
  t.is(
    code(() => book.addTable(0, 'B2:C5')),
    'TABLE_OVERLAP'
  )
  t.is(
    code(() => book.addTable(0, 'D1:E3')),
    'RANGE_HAS_MERGES'
  )
  t.is(
    code(() => book.addTable(0, 'G1:H3', { name: 'table1' })),
    'INVALID_DEFINED_NAME'
  )
  book.defineName('Prices', 'Sheet1!$A$1')
  t.is(
    code(() => book.addTable(0, 'G1:H3', { name: 'Prices' })),
    'INVALID_DEFINED_NAME'
  )
  t.is(
    code(() => book.defineName('Table1', 'Sheet1!$A$1')),
    'INVALID_DEFINED_NAME'
  )
  t.is(
    code(() => book.addTable(0, 'G1:H1')),
    'INVALID_RANGE'
  )
  book.setCells(0, { G4: 'busy' })
  t.is(
    code(() => book.addTable(0, 'G1:H3', { totals: true })),
    'INVALID_RANGE'
  )
  book.autofilter(0, 'J1:K5')
  t.is(
    code(() => book.addTable(0, 'K1:L3')),
    'TABLE_OVERLAP'
  )
  t.is(
    code(() => book.merge(0, 'A2:B2')),
    'TABLE_OVERLAP'
  )
})

test('deleteTable turns structured references into ranges', (t) => {
  const book = xlsx.create({ sheets: ['Data', 'Report'] })
  book.appendRows('Data', [
    ['Item', 'Cost'],
    ['Rent', 100],
    ['Food', 50]
  ])
  book.addTable('Data', 'A1:B3', { name: 'Costs', totals: true })
  book.setCells('Report', {
    A1: '=SUM(Costs[Cost])',
    A2: '=ROWS(Costs[#All])',
    A3: '=Costs[[#Headers],[Item]]'
  })
  t.alike(book.deleteTable('costs'), { deleted: 'Costs', converted: 4 })
  const again = reopen(book)
  t.alike(again.tables(), [])
  t.alike(again.read('Data').formulas, { B4: '=SUBTOTAL(109,Data!$B$2:$B$3)' })
  t.alike(again.read('Report').formulas, {
    A1: '=SUM(Data!$B$2:$B$3)',
    A2: '=ROWS(Data!$A$1:$B$4)',
    A3: '=Data!$A$1'
  })
  const files = unzip(again.save())
  t.absent(files['xl/tables/table1.xml'])
  t.absent(decode(files['xl/worksheets/sheet1.xml']).includes('tablePart'))
  t.is(
    code(() => again.deleteTable('Costs')),
    'TABLE_NOT_FOUND'
  )
})

test('deleteTable refuses references it cannot rewrite', (t) => {
  const book = xlsx.create()
  book.appendRows(0, [['Cost'], [1], [2]])
  book.addTable(0, 'A1:A3', { name: 'T' })
  book.setCells(0, { C2: '=T[@Cost]*2' })
  t.is(
    code(() => book.deleteTable('T')),
    'TABLE_IS_REFERENCED'
  )
  t.is(book.tables().length, 1)
})

test('tables in an existing workbook are listed', (t) => {
  const book = xlsx.open(corpusFile('xlsx/ExcelTables.xlsx'))
  const tables = book.tables()
  t.ok(tables.length > 0)
  t.ok(tables.every((table) => table.columns.length > 0 && table.range.includes(':')))
})
