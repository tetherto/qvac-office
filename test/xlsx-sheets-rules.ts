import test from 'brittle'
import * as xlsx from '../lib/xlsx/index.ts'
import { DocsError } from '../lib/errors.ts'
import { decode, unzip } from './helpers/engine.ts'

const reopen = (book: xlsx.Workbook): xlsx.Workbook => xlsx.open(book.save())
const partXml = (book: xlsx.Workbook, name: string): string => decode(unzip(book.save())[name])
const sheetXml = (book: xlsx.Workbook): string => partXml(book, 'xl/worksheets/sheet1.xml')

function code(fn: () => unknown): string | undefined {
  try {
    fn()
  } catch (err) {
    return err instanceof DocsError ? err.code : 'not a DocsError'
  }

  return undefined
}

test('conditional formats of every kind round trip', (t) => {
  const book = xlsx.create()
  book.setCells(0, { A1: 5, A2: 50 })
  const red = { fill: 'FFC7CE', color: '9C0006', bold: true }
  t.alike(
    book.addConditionalFormat(0, 'A1:A10', {
      kind: 'cellIs',
      operator: 'greaterThan',
      value: 10,
      style: red
    }),
    { range: 'A1:A10', kind: 'cellIs', priority: 1 }
  )
  book.addConditionalFormat(0, 'A1:A10', {
    kind: 'cellIs',
    operator: 'between',
    values: [1, '=$C$1'],
    style: red
  })
  book.addConditionalFormat(0, 'B1:B10', {
    kind: 'colorScale',
    colors: ['F8696B', 'FFEB84', '63BE7B']
  })
  book.addConditionalFormat(0, 'C1:C10 E1:E10', { kind: 'dataBar', color: '638EC6' })
  book.addConditionalFormat(0, 'D2:D10', {
    kind: 'text',
    contains: 'late',
    style: { italic: true }
  })
  const again = reopen(book)
  const sheet = sheetXml(again)
  t.ok(
    sheet.includes(
      '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>10</formula></cfRule>'
    )
  )
  t.ok(sheet.includes('operator="between"><formula>1</formula><formula>$C$1</formula>'))
  t.ok(sheet.includes('<cfvo type="percentile" val="50"/>'))
  t.ok(sheet.includes('<conditionalFormatting sqref="C1:C10 E1:E10">'))
  t.ok(sheet.includes('<color rgb="FF638EC6"/>'))
  t.ok(sheet.includes('text="late"><formula>NOT(ISERROR(SEARCH("late",D2)))</formula>'))
  const styles = partXml(again, 'xl/styles.xml')
  t.ok(
    styles.includes(
      '<dxfs count="2"><dxf><font><b val="1"/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>'
    )
  )
  t.alike(again.deleteConditionalFormats(0, 'A5'), { deleted: 2 })
  t.alike(again.deleteConditionalFormats(0), { deleted: 3 })
  t.absent(sheetXml(again).includes('conditionalFormatting'))
})

test('conditional format rules are checked', (t) => {
  const book = xlsx.create()
  t.is(
    code(() =>
      book.addConditionalFormat(0, 'A1', {
        kind: 'cellIs',
        operator: 'between',
        value: 1,
        style: { bold: true }
      })
    ),
    'INVALID_OPTIONS'
  )
  t.is(
    code(() => book.addConditionalFormat(0, 'A1', { kind: 'colorScale', colors: ['FF0000'] })),
    'INVALID_OPTIONS'
  )
  t.is(
    code(() => book.addConditionalFormat(0, 'A1', { kind: 'dataBar', color: 'blue' })),
    'INVALID_STYLE'
  )
  t.is(
    code(() =>
      book.addConditionalFormat(0, 'A1', {
        kind: 'text',
        contains: 'x',
        style: { size: 3 } as xlsx.DiffStyle
      })
    ),
    'INVALID_STYLE'
  )
  t.is(
    code(() => book.addConditionalFormat(0, 'nope', { kind: 'dataBar', color: '000000' })),
    'INVALID_CELL'
  )
})

test('data validations of every type round trip', (t) => {
  const book = xlsx.create()
  t.alike(
    book.addValidation(0, 'A2:A20', { type: 'list', values: ['Yes', 'No'], prompt: 'Pick one' }),
    {
      range: 'A2:A20',
      type: 'list'
    }
  )
  book.addValidation(0, 'B2:B20', {
    type: 'whole',
    min: 0,
    max: 10,
    error: { title: 'Out of range', text: '0 to 10', style: 'warning' }
  })
  book.addValidation(0, 'C2:C20', { type: 'decimal', min: 0.5 })
  book.addValidation(0, 'D2:D20', { type: 'date', max: '2026-12-31', allowBlank: false })
  t.alike(book.addValidation(0, 'E2', { type: 'textLength', operator: 'lessThan', max: 20 }), {
    range: 'E2',
    type: 'textLength',
    operator: 'lessThan'
  })
  book.addValidation(0, 'F2:F9', { type: 'list', values: '=$H$1:$H$5' })
  const again = reopen(book)
  const sheet = sheetXml(again)
  t.ok(sheet.includes('<dataValidations count="6">'))
  t.ok(
    sheet.includes(
      'type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" prompt="Pick one" sqref="A2:A20"><formula1>"Yes,No"</formula1>'
    )
  )
  t.ok(sheet.includes('errorStyle="warning" operator="between"'))
  t.ok(sheet.includes('<formula1>0</formula1><formula2>10</formula2>'))
  t.ok(
    sheet.includes(
      'operator="greaterThanOrEqual" allowBlank="1" showErrorMessage="1" sqref="C2:C20"><formula1>0.5</formula1>'
    )
  )
  t.ok(
    sheet.includes(
      'type="date" operator="lessThanOrEqual" showErrorMessage="1" sqref="D2:D20"><formula1>46387</formula1>'
    )
  )
  t.ok(sheet.includes('<formula1>$H$1:$H$5</formula1>'))
  t.alike(again.deleteValidations(0, 'A1:B2'), { deleted: 2 })
  t.alike(again.deleteValidations(0), { deleted: 4 })
  t.absent(sheetXml(again).includes('dataValidations'))
})

test('data validations are checked and do not stack', (t) => {
  const book = xlsx.create()
  book.addValidation(0, 'A1:A5', { type: 'whole', min: 1 })
  t.is(
    code(() => book.addValidation(0, 'A3', { type: 'decimal', min: 1 })),
    'VALIDATION_OVERLAP'
  )
  t.is(
    code(() => book.addValidation(0, 'B1', { type: 'whole', min: 1.5 })),
    'INVALID_OPTIONS'
  )
  t.is(
    code(() => book.addValidation(0, 'B1', { type: 'whole', operator: 'between', min: 1 })),
    'INVALID_OPTIONS'
  )
  t.is(
    code(() => book.addValidation(0, 'B1', { type: 'list', values: ['a,b'] })),
    'INVALID_OPTIONS'
  )
  t.is(
    code(() => book.addValidation(0, 'B1', { type: 'time' } as unknown as xlsx.ValidationRule)),
    'INVALID_OPTIONS'
  )
  t.is(
    code(() => book.addValidation(0, 'B1', { type: 'date', min: '2026-02-30' })),
    'INVALID_VALUE'
  )
})
