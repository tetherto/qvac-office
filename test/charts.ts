import test from 'brittle'
import * as pptx from '../lib/pptx/index.ts'
import { Package } from '../lib/opc/package.ts'
import * as units from '../lib/units.ts'
import * as xml from '../lib/xml.ts'
import * as utf8 from '../lib/utf8.ts'
import { fixture, decode, unzip, zipFiles } from './helpers/engine.ts'
import type { ChartInfo } from '../lib/drawingml/chart.ts'
const open = (): pptx.Presentation => pptx.open(fixture('charts.pptx'))
const reopen = (deck: pptx.Presentation): pptx.Presentation => pptx.open(deck.save())

function chartOf(deck: pptx.Presentation, slide: number): ChartInfo {
  const found = deck.read()[slide].shapes.find((s) => s.kind === 'chart')?.chart
  if (!found) {
    throw new Error(`no chart on slide ${slide}`)
  }

  return found
}

function bad<T>(value: unknown): T {
  return value as T
}

function workbookRows(bytes: Uint8Array): (string | number | null)[][] {
  const pkg = Package.open(bytes)
  const sheet = pkg.names().find((name) => name.includes('worksheets/sheet'))
  const doc = pkg.xml(sheet ?? '')
  return Array.from(doc.getElementsByTagName('row')).map((row) =>
    Array.from(row.getElementsByTagName('c')).map((c) => {
      const v = c.getElementsByTagName('v')[0]
      const t = c.getElementsByTagName('t')[0]
      return v ? Number(v.textContent) : (t?.textContent ?? null)
    })
  )
}

test('reads chart type, title, categories and series', (t) => {
  const deck = open()
  t.alike(chartOf(deck, 0), {
    type: 'column',
    title: 'Revenue by quarter',
    categories: ['Q1', 'Q2', 'Q3'],
    series: [
      { name: '2025', values: [10, 12.5, 14] },
      { name: '2026', values: [11, 13, 15.5] }
    ]
  })
  t.is(chartOf(deck, 1).type, 'pie')
})

test('setChartData rewrites the caches and the embedded workbook', (t) => {
  const deck = open()
  const data = {
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series: [
      { name: '2025', values: [10, 12.5, 14, 16] },
      { name: '2026', values: [11, 13, 15.5, null] }
    ]
  }
  t.alike(deck.setChartData({ slide: 0, shape: 3 }, data), { series: 2, categories: 4 })

  const again = reopen(deck)
  const read = chartOf(again, 0)
  t.alike(read.categories, data.categories)
  t.alike(read.series, [
    { name: '2025', values: [10, 12.5, 14, 16] },
    { name: '2026', values: [11, 13, 15.5, null] }
  ])
  t.is(read.title, 'Revenue by quarter')

  const files = unzip(again.save())
  const chartXml = decode(files['ppt/charts/chart1.xml'])
  t.ok(chartXml.includes('<c:f>Sheet1!$A$2:$A$5</c:f>'))
  t.ok(chartXml.includes('<c:f>Sheet1!$C$2:$C$5</c:f>'))
  t.alike(workbookRows(files['ppt/embeddings/Microsoft_Excel_Sheet1.xlsx']), [
    ['2025', '2026'],
    ['Q1', 10, 11],
    ['Q2', 12.5, 13],
    ['Q3', 14, 15.5],
    ['Q4', 16]
  ])
})

test('setChartData adds and removes series', (t) => {
  const deck = open()
  deck.setChartData(
    { slide: 0, shape: 3 },
    {
      categories: ['A', 'B'],
      series: [
        { name: 'one', values: [1, 2] },
        { name: 'two', values: [3, 4] },
        { name: 'three', values: [5, 6] }
      ]
    }
  )
  let read = chartOf(reopen(deck), 0)
  t.alike(
    read.series.map((s) => s.name),
    ['one', 'two', 'three']
  )
  const xml = decode(unzip(deck.save())['ppt/charts/chart1.xml'])
  t.ok(xml.includes('<c:idx val="2"/><c:order val="2"/>'))

  deck.setChartData(
    { slide: 0, shape: 3 },
    { categories: ['A'], series: [{ name: 'only', values: [9] }] }
  )
  read = chartOf(reopen(deck), 0)
  t.alike(read.series, [{ name: 'only', values: [9] }])
})

test('setChartData takes numeric categories', (t) => {
  const deck = open()
  deck.setChartData(
    { slide: 1, shape: 2 },
    { categories: [2024, 2025], series: [{ name: 'Share', values: [1, 2] }] }
  )
  t.alike(chartOf(reopen(deck), 1).categories, [2024, 2025])
})

test('setChartData keeps the number format of date categories', (t) => {
  const source = unzip(fixture('charts.pptx'))
  const dates =
    '<c:cat><c:numRef><c:f>Sheet1!$A$2:$A$4</c:f><c:numCache><c:formatCode>m/d/yyyy</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>45292</c:v></c:pt><c:pt idx="1"><c:v>45323</c:v></c:pt><c:pt idx="2"><c:v>45352</c:v></c:pt></c:numCache></c:numRef></c:cat>'
  source['ppt/charts/chart1.xml'] = utf8.encode(
    decode(source['ppt/charts/chart1.xml']).replace(/<c:cat>.*?<\/c:cat>/g, dates)
  )
  const deck = pptx.open(zipFiles(source))
  const read = chartOf(deck, 0)
  t.alike(read.categories, [45292, 45323, 45352])
  deck.setChartData(
    { slide: 0, shape: 3 },
    {
      categories: [45292, 45323, 45352],
      series: [
        { name: '2025', values: [10, 12.5, 14] },
        { name: '2026', values: [11, 13, 15.5] }
      ]
    }
  )
  const chartXml = decode(unzip(deck.save())['ppt/charts/chart1.xml'])
  const cat = chartXml.match(/<c:cat>.*?<\/c:cat>/)?.[0] ?? ''
  t.ok(cat.includes('<c:formatCode>m/d/yyyy</c:formatCode>'))
})

test('setChartData leaves characters XML cannot hold out of the chart', (t) => {
  const deck = open()
  deck.setChartData(
    { slide: 0, shape: 3 },
    { categories: ['Q\u00011'], series: [{ name: 'a\u000bb', values: [1] }] }
  )
  const chartXml = decode(unzip(deck.save())['ppt/charts/chart1.xml'])
  t.absent(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(chartXml))
  t.ok(xml.parse(chartXml))
  const read = chartOf(reopen(deck), 0)
  t.alike(read.categories, ['Q1'])
  t.is(read.series[0].name, 'ab')
})

test('setChartData validates its input', (t) => {
  const deck = open()
  t.exception(
    () =>
      deck.setChartData(
        { slide: 0, shape: 3 },
        { categories: ['a', 'b'], series: [{ name: 'x', values: [1] }] }
      ),
    /series 0 \("x"\) has 1 values but there are 2 categories/
  )
  t.exception(
    () => deck.setChartData({ slide: 0, shape: 3 }, { categories: [], series: [] }),
    /non-empty/
  )
  t.exception(
    () =>
      deck.setChartData(
        { slide: 0, shape: 3 },
        { categories: ['a'], series: [bad({ name: 'x', values: ['1'] })] }
      ),
    /must be a number or null/
  )
  t.exception(
    () => deck.setChartData({ slide: 0, shape: 2 }, { categories: ['a'], series: [] }),
    /is a text, not a chart/
  )
  t.is(deck.changed, false)
})

test('setChartTitle keeps the title formatting, adds and removes titles', (t) => {
  const deck = open()
  deck.setChartTitle({ slide: 0, shape: 3 }, 'Revenue, restated')
  let xml = decode(unzip(deck.save())['ppt/charts/chart1.xml'])
  t.ok(xml.includes('<a:rPr b="1"/><a:t>Revenue, restated</a:t>'))

  deck.setChartTitle({ slide: 1, shape: 2 }, 'Market share')
  t.is(chartOf(reopen(deck), 1).title, 'Market share')

  deck.setChartTitle({ slide: 0, shape: 3 }, null)
  t.is(chartOf(reopen(deck), 0).title, null)
  xml = decode(unzip(deck.save())['ppt/charts/chart1.xml'])
  t.ok(xml.includes('<c:autoTitleDeleted val="1"/>'))
})

test('addChart builds each supported type with an embedded workbook', (t) => {
  for (const type of ['column', 'bar', 'line', 'pie', 'doughnut', 'area'] as const) {
    const deck = pptx.create()
    deck.addSlide({ layout: 'Blank' })
    const placed = deck.addChart(0, {
      type,
      title: `A ${type} chart`,
      categories: ['North', 'South', 'East'],
      series: [
        { name: 'Units', values: [3, 5, 2] },
        { name: 'Returns', values: [1, 0, 1] }
      ],
      x: units.inches(1),
      y: units.inches(1),
      width: units.inches(6),
      height: units.inches(4)
    })
    t.is(placed.width, units.inches(6), type)

    const again = reopen(deck)
    const read = chartOf(again, 0)
    t.is(read.type, type)
    t.is(read.title, `A ${type} chart`)
    t.alike(read.series[1], { name: 'Returns', values: [1, 0, 1] })

    const files = unzip(again.save())
    t.ok(files['ppt/charts/chart1.xml'])
    t.ok(files['ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'])
    t.ok(decode(files['[Content_Types].xml']).includes('/ppt/charts/chart1.xml'))
    t.ok(decode(files['[Content_Types].xml']).includes('Extension="xlsx"'))
  }
})

test('addChart validates the type and options', (t) => {
  const deck = open()
  const data = { categories: ['a'], series: [{ name: 'x', values: [1] }] }
  t.exception(
    () => deck.addChart(0, bad({ type: 'radar', ...data })),
    /chart type must be one of column/
  )
  t.exception(
    () => deck.addChart(0, bad({ type: 'column', stacked: 'yes', ...data })),
    /stacked must be/
  )
  t.exception(
    () => deck.addChart(0, bad({ type: 'column', legend: 'middle', ...data })),
    /legend must be/
  )
})

test('addChart stacks and hides the legend', (t) => {
  const deck = open()
  deck.addChart(0, {
    type: 'column',
    stacked: 'percent',
    legend: false,
    categories: ['a', 'b'],
    series: [
      { name: 'x', values: [1, 2] },
      { name: 'y', values: [3, 4] }
    ]
  })
  const xml = decode(unzip(deck.save())['ppt/charts/chart3.xml'])
  t.ok(xml.includes('<c:grouping val="percentStacked"/>'))
  t.ok(xml.includes('<c:overlap val="100"/>'))
  t.absent(xml.includes('<c:legend>'))
})

test('deleting a slide drops its chart and workbook', (t) => {
  const deck = open()
  deck.deleteSlide(0)
  const files = unzip(deck.save())
  t.absent(files['ppt/charts/chart1.xml'])
  t.absent(files['ppt/embeddings/Microsoft_Excel_Sheet1.xlsx'])
  t.ok(files['ppt/charts/chart2.xml'])
})
