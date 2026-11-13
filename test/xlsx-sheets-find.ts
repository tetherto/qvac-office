import test from 'brittle'
import * as xlsx from '../lib/xlsx/index.ts'
import { DocsError } from '../lib/errors.ts'
import { corpusFile, fixture } from './helpers/engine.ts'

function code(fn: () => unknown): string | undefined {
  try {
    fn()
  } catch (err) {
    return err instanceof DocsError ? err.code : 'not a DocsError'
  }

  return undefined
}

function book(): xlsx.Workbook {
  const wb = xlsx.create({ sheets: ['Data', 'Notes'] })
  wb.appendRows('Data', [
    ['Region', 'Units', 'Total'],
    ['North', 3, '=B2*2'],
    ['South', 5, '=B3*2']
  ])
  wb.setCells('Notes', { A1: 'north is ahead', B4: true })

  return xlsx.open(wb.save())
}

test('find searches values, optionally formulas, case and regex aware', (t) => {
  const wb = book()
  t.alike(wb.find('north'), [
    { sheet: 'Data', ref: 'A2', value: 'North' },
    { sheet: 'Notes', ref: 'A1', value: 'north is ahead' }
  ])
  t.alike(wb.find('North', { matchCase: true }), [{ sheet: 'Data', ref: 'A2', value: 'North' }])
  t.alike(wb.find('north', { sheet: 'Notes' }).length, 1)
  t.alike(wb.find('^S', { regex: true }), [{ sheet: 'Data', ref: 'A3', value: 'South' }])
  t.alike(wb.find('B3', { formulas: true }), [
    { sheet: 'Data', ref: 'C3', value: 10, formula: '=B3*2' }
  ])
  t.alike(wb.find('true'), [{ sheet: 'Notes', ref: 'B4', value: true }])
  t.is(wb.find('o', { limit: 2 }).length, 2)
  t.is(
    code(() => wb.find('(', { regex: true })),
    'INVALID_OPTIONS'
  )
  t.is(
    code(() => wb.find('')),
    'INVALID_OPTIONS'
  )
  t.is(
    code(() => wb.find('x', { sheet: 'Missing' })),
    'SHEET_NOT_FOUND'
  )
})

test('outline summarises every sheet', (t) => {
  const wb = book()
  wb.addTable('Data', 'A1:C3')
  wb.addChart('Notes', {
    type: 'column',
    categories: 'Data!A2:A3',
    series: [{ values: 'Data!B2:B3' }]
  })
  t.alike(wb.outline(), [
    {
      index: 0,
      name: 'Data',
      kind: 'worksheet',
      state: 'visible',
      range: 'A1:C3',
      header: { row: 1, names: ['Region', 'Units', 'Total'] },
      rows: 3,
      columns: 3,
      charts: 0,
      tables: 1
    },
    {
      index: 1,
      name: 'Notes',
      kind: 'worksheet',
      state: 'visible',
      range: 'A1:B4',
      header: null,
      rows: 4,
      columns: 2,
      charts: 1,
      tables: 0
    }
  ])
})

test('outline stays small for chart sheets and real files', (t) => {
  const sheets = xlsx.open(corpusFile('xlsx/WithChartSheet.xlsx')).outline()
  t.ok(sheets.some((s) => s.kind !== 'worksheet' && s.range === null))
  const sales = xlsx.open(fixture('sales.xlsx')).outline()
  t.ok(sales.length > 0)
  t.ok(JSON.stringify(sales).length < 2000)
})
