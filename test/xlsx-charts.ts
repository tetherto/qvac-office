import test from 'brittle'
import * as xlsx from '../lib/xlsx/index.ts'
import * as units from '../lib/units.ts'
import * as xml from '../lib/xml.ts'
import { fixture, decode, unzip } from './helpers/engine.ts'

const X = xml.ns.x
const reopen = (book: xlsx.Workbook): xlsx.Workbook => xlsx.open(book.save())
const chartXml = (book: xlsx.Workbook): string => decode(unzip(book.save())['xl/charts/chart1.xml'])

function unchecked(value: unknown): value is never {
  return value !== undefined
}

function withSharedFormula(): xlsx.Workbook {
  const book = xlsx.create()
  book.setCells(0, { A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, B3: 30 })
  const doc = book.package.xml('/xl/worksheets/sheet1.xml')
  const rows = xml.descendants(doc, X, 'row')
  for (let i = 0; i < 3; i++) {
    const c = xml.named(doc, 'c', X)
    c.setAttribute('r', `C${i + 1}`)
    const f = xml.named(doc, 'f', X)
    f.setAttribute('t', 'shared')
    f.setAttribute('si', '0')
    if (i === 0) {
      f.setAttribute('ref', 'C1:C3')
      f.appendChild(doc.createTextNode('A1*B1+$B$1'))
    }

    c.appendChild(f)
    rows[i].appendChild(c)
  }

  book.package.touch('/xl/worksheets/sheet1.xml')

  return reopen(book)
}

test('shared formulas read as each cell sees them', (t) => {
  const book = withSharedFormula()
  t.alike(book.read(0).formulas, { C1: '=A1*B1+$B$1', C2: '=A2*B2+$B$1', C3: '=A3*B3+$B$1' })
  t.is(book.get(0, 'C3').formula, '=A3*B3+$B$1')
  t.exception(() => book.setCells(0, { C1: 5 }), /part of a shared formula over C1:C3/)
})

test('addChart plots cell ranges without an embedded workbook', (t) => {
  const book = xlsx.open(fixture('sales.xlsx'))
  const placed = book.addChart('Sales', {
    type: 'column',
    title: 'Units by quarter',
    categories: 'A3:A5',
    series: [
      { values: 'B3:B5', nameCell: 'B2' },
      { values: 'C3:C5', name: 'Price' }
    ]
  })
  t.alike(placed, { cell: 'J1', width: 5486400, height: 3200400 })
  const again = reopen(book)
  const charts = again.read('Sales').charts ?? []
  t.is(charts.length, 1)
  t.alike(charts[0], {
    type: 'column',
    title: 'Units by quarter',
    categories: ['Q1', 'Q2', 'Q3'],
    series: [
      { name: 'Units', values: [120, 150, 90] },
      { name: 'Price', values: [9.5, 9.75, 10] }
    ],
    cell: 'J1'
  })
  const files = unzip(again.save())
  const chart = decode(files['xl/charts/chart1.xml'])
  t.ok(chart.includes('<c:f>Sales!$A$3:$A$5</c:f>'))
  t.ok(chart.includes('<c:f>Sales!$B$2</c:f>'))
  t.ok(chart.includes('<c:tx><c:v>Price</c:v></c:tx>'))
  t.absent(chart.includes('externalData'))
  t.absent(Object.keys(files).some((n) => n.includes('embeddings')))
  t.ok(decode(files['xl/drawings/drawing1.xml']).includes('<xdr:graphicFrame macro="">'))
})

test('addChart reads ranges on other sheets and validates them', (t) => {
  const book = xlsx.open(fixture('sales.xlsx'))
  book.addSheet('Chart')
  book.addChart('Chart', {
    type: 'line',
    categories: 'Sales!A3:A5',
    series: [{ values: "'Sales'!B3:B5", name: 'Units' }],
    cell: 'B2',
    width: units.inches(4)
  })
  const charts = reopen(book).read('Chart').charts ?? []
  t.alike(charts[0].series[0].values, [120, 150, 90])
  t.exception(
    () =>
      book.addChart('Sales', {
        type: 'column',
        categories: 'A3:B5',
        series: [{ values: 'B3:B5' }]
      }),
    /single row or a single column/
  )
  t.exception(
    () =>
      book.addChart('Sales', {
        type: 'column',
        categories: 'A3:A5',
        series: [{ values: 'B3:B4' }]
      }),
    /covers 2 cells but the categories cover 3/
  )
  const radar: unknown = { type: 'radar', categories: 'A3:A5', series: [{ values: 'B3:B5' }] }
  if (unchecked(radar)) {
    t.exception(() => book.addChart('Sales', radar), /chart type must be one of/)
  }
})

test('charts follow sheet renames and row shifts and go with their sheet', (t) => {
  const book = xlsx.create({ sheets: ['Data', 'Other'] })
  book.appendRows('Data', [
    ['a', 1],
    ['b', 2]
  ])
  book.addChart('Data', { type: 'pie', categories: 'A1:A2', series: [{ values: 'B1:B2' }] })
  book.renameSheet('Data', 'Figures')
  t.ok(chartXml(book).includes('<c:f>Figures!$B$1:$B$2</c:f>'))
  book.insertRows('Figures', { at: 1, count: 1 })
  t.ok(chartXml(book).includes('<c:f>Figures!$B$2:$B$3</c:f>'))
  t.execution(() => book.deleteSheet('Figures'))
  const files = unzip(book.save())
  t.absent(files['xl/charts/chart1.xml'])
  t.absent(files['xl/drawings/drawing1.xml'])
})

test('a chart on another sheet keeps its source sheet from being deleted', (t) => {
  const book = xlsx.create({ sheets: ['Data', 'Dash'] })
  book.appendRows('Data', [['a', 1]])
  book.addChart('Dash', {
    type: 'bar',
    categories: 'Data!A1:A1',
    series: [{ values: 'Data!B1:B1' }]
  })
  t.exception(() => book.deleteSheet('Data'), /is used by 1 chart\(s\)/)
})
