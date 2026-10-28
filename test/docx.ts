import test from 'brittle'
import * as docx from '../lib/docx/index.ts'
import * as units from '../lib/units.ts'
import { fixture, decode, unzip } from './helpers/engine.ts'
import type { Block, Document, ParagraphBlock, TableBlock } from '../lib/docx/index.ts'

const open = (): Document => docx.open(fixture('report.docx'))
const reopen = (doc: Document): Document => docx.open(doc.save())
const blocks = (doc: Document): Block[] => doc.read().blocks
const texts = (doc: Document): string[] =>
  blocks(doc).map((b) => (b.type === 'table' ? '[table]' : b.text))
const documentXml = (doc: Document): string => decode(unzip(doc.save())['word/document.xml'])

function para(block: Block | undefined): ParagraphBlock {
  if (!block || block.type !== 'paragraph') {
    throw new Error('expected a paragraph block')
  }

  return block
}

function table(block: Block | undefined): TableBlock {
  if (!block || block.type !== 'table') {
    throw new Error('expected a table block')
  }

  return block
}

test('reads paragraphs, styles, headings, lists, tables, images and headers', (t) => {
  const read = open().read()
  t.is(read.blocks.length, 14)
  t.alike(read.blocks[0], {
    index: 0,
    type: 'paragraph',
    style: 'Title',
    text: 'Quarterly Report',
    heading: 0
  })
  t.alike(para(read.blocks[3]).text, 'Revenue grew twelve percent year over year.')
  t.is(para(read.blocks[4]).style, 'List Bullet')
  t.ok(para(read.blocks[4]).list)
  t.is(para(read.blocks[6]).heading, 2)
  t.alike(table(read.blocks[10]).rows, [
    ['Metric', '2025', '2026'],
    ['Margin', '25%', '36%']
  ])
  t.is(para(read.blocks[13]).images, 1)
  t.alike(read.headers, ['Confidential'])
})

test('rejects files that are not Word documents', (t) => {
  t.exception(() => docx.open(fixture('basic.pptx')), /not a Word document/)
})

test('setText keeps the paragraph style and the first run formatting', (t) => {
  const doc = open()
  t.alike(doc.setText(3, 'Revenue fell two percent.'), { changed: true })
  t.alike(doc.setText(4, 'North led the growth'), { changed: false })
  const again = reopen(doc)
  t.is(para(blocks(again)[3]).text, 'Revenue fell two percent.')
  t.is(para(blocks(again)[4]).style, 'List Bullet')
  t.ok(documentXml(again).includes('<w:r><w:t>Revenue fell two percent.</w:t></w:r>'))

  doc.setText(2, 'Income')
  t.is(para(blocks(reopen(doc))[2]).style, 'Heading 1')
})

test('setText takes formatted runs, line breaks and tabs', (t) => {
  const doc = open()
  doc.setText(1, {
    runs: ['Prepared by ', { text: 'finance', bold: true, color: 'C0392B' }, '\nand audit\tteams']
  })
  const xml = documentXml(doc)
  t.ok(xml.includes('<w:t xml:space="preserve">Prepared by </w:t>'))
  t.ok(xml.includes('<w:rPr><w:b/><w:bCs/><w:color w:val="C0392B"/></w:rPr><w:t>finance</w:t>'))
  t.ok(xml.includes('<w:br/><w:t>and audit</w:t><w:tab/><w:t>teams</w:t>'))
  t.is(para(blocks(reopen(doc))[1]).text, 'Prepared by finance\nand audit\tteams')
})

test('errors name what exists', (t) => {
  const doc = open()
  t.exception(
    () => doc.setText(40, 'x'),
    /block 40 does not exist; the document has 14 blocks \(0 to 13\)/
  )
  t.exception(() => doc.setText(10, 'x'), /block 10 is a table; use setCell/)
  t.exception(
    () => doc.insertParagraphs(1, [{ text: 'x', style: 'Fancy' }]),
    /no paragraph style "Fancy"; it has "Normal"/
  )
  t.exception(
    () => doc.replaceSection(3, { body: ['x'] }),
    /block 3 is not a heading; the headings are 2 "Revenue"/
  )
  t.is(doc.changed, false)
})

test('insertParagraphs adds styled paragraphs in place', (t) => {
  const doc = open()
  const result = doc.insertParagraphs(6, [
    { text: 'East opened', style: 'List Bullet' },
    'West grew too',
    { pageBreak: true }
  ])
  t.alike(result, { index: 6, count: 3 })
  const read = blocks(reopen(doc))
  t.is(para(read[6]).style, 'List Bullet')
  t.is(para(read[7]).text, 'West grew too')
  t.is(para(read[7]).style, 'List Bullet')
  t.ok(para(read[8]).pageBreak)
  t.is(para(read[9]).text, 'Regions')
})

test('insertParagraphs after a heading does not inherit the heading style', (t) => {
  const doc = open()
  doc.insertParagraphs(3, 'Summary first.')
  t.is(para(blocks(reopen(doc))[3]).style, 'Normal')
})

test('appendParagraphs goes before the section properties', (t) => {
  const doc = open()
  doc.appendParagraphs([{ text: 'Appendix', style: 'Heading 1' }, 'Details.'])
  const read = blocks(reopen(doc))
  t.alike(
    read.slice(-2).map((b) => [para(b).text, para(b).style]),
    [
      ['Appendix', 'Heading 1'],
      ['Details.', 'Normal']
    ]
  )
  t.ok(documentXml(doc).trim().endsWith('</w:sectPr></w:body></w:document>'))
})

test('replaceSection swaps a heading and everything under it', (t) => {
  const doc = open()
  const result = doc.replaceSection(2, {
    title: 'Income',
    body: ['Income rose.', 'It rose everywhere.']
  })
  t.alike(result, { index: 2, removed: 5, added: 2 })
  t.alike(texts(reopen(doc)).slice(0, 7), [
    'Quarterly Report',
    'Prepared by the finance team.',
    'Income',
    'Income rose.',
    'It rose everywhere.',
    'Costs',
    'Costs fell as contracts ended.'
  ])
})

test('replaceSection removes tables inside the section', (t) => {
  const doc = open()
  doc.replaceSection(8, { body: ['Costs are no longer tracked.'] })
  t.alike(texts(reopen(doc)).slice(8, 11), ['Costs', 'Costs are no longer tracked.', 'Outlook'])
})

test('a subsection ends at the next heading of its rank or higher', (t) => {
  const doc = open()
  doc.replaceSection(6, { body: ['All regions reported on time.'] })
  t.alike(texts(reopen(doc)).slice(6, 9), ['Regions', 'All regions reported on time.', 'Costs'])
})

test('deleteSection and deleteBlocks remove content', (t) => {
  const doc = open()
  doc.deleteSection(11)
  const after = blocks(doc)
  t.is(after.length, 12)
  t.is(after[10].type, 'table')
  t.alike(after[11], { index: 11, type: 'paragraph', style: 'Normal', text: '' })
  doc.deleteBlocks(0, { count: 2 })
  t.is(para(blocks(reopen(doc))[0]).text, 'Revenue')
  t.exception(() => doc.deleteBlocks(0, { count: 99 }), /count must be from 1 to/)
})

test('deleting the image paragraph drops the media part', (t) => {
  const doc = open()
  doc.deleteBlocks(13)
  const files = unzip(doc.save())
  t.absent(Object.keys(files).some((name) => name.startsWith('word/media/')))
  t.absent(decode(files['word/_rels/document.xml.rels']).includes('relationships/image'))
})

test('formatParagraph and formatStyle change formatting', (t) => {
  const doc = open()
  doc.formatParagraph(3, { size: 14, italic: true, align: 'center' })
  let xml = documentXml(doc)
  t.ok(xml.includes('<w:jc w:val="center"/>'))
  t.ok(
    xml.includes(
      '<w:rPr><w:b/><w:i/><w:iCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>twelve percent</w:t>'
    )
  )

  t.alike(doc.formatStyle('normal', { size: 14, font: 'Georgia' }), { style: 'Normal' })
  xml = decode(unzip(doc.save())['word/styles.xml'])
  t.ok(
    xml.includes(
      '<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:cs="Georgia" w:eastAsia="Georgia"/>'
    )
  )
  t.exception(() => doc.formatStyle('Normal', { size: 0 }), /size is in points/)
  t.exception(() => doc.formatParagraph(3, { shout: true } as never), /unknown format "shout"/)
})

test('tables take cells, rows and new tables', (t) => {
  const doc = open()
  doc.setCell({ table: 10, row: 1, column: 2 }, '38%')
  doc.addRow(10, ['Headcount', 120, 115])
  doc.addRow(10, ['Revenue', '$1m', '$1.2m'], { at: 1 })
  doc.deleteRow({ table: 10, row: 3 })
  t.alike(table(blocks(reopen(doc))[10]).rows, [
    ['Metric', '2025', '2026'],
    ['Revenue', '$1m', '$1.2m'],
    ['Margin', '25%', '38%']
  ])
  t.exception(
    () => doc.setCell({ table: 10, row: 0, column: 5 }, 'x'),
    /column 5 does not exist; the row has 3/
  )
  t.exception(() => doc.addRow(10, [1, 2, 3, 4]), /the row has 3 cells but 4 values/)

  const added = doc.addTable(12, {
    rows: [
      ['Region', 'Units'],
      ['North', 3]
    ]
  })
  t.alike(added, { index: 12, rows: 2, columns: 2 })
  const read = table(blocks(reopen(doc))[12])
  t.alike(read.rows, [
    ['Region', 'Units'],
    ['North', '3']
  ])
  t.ok(documentXml(doc).includes('<w:tblStyle w:val="TableGrid"/>'))
})

test('addImage inserts an inline picture sized by aspect ratio', (t) => {
  const doc = open()
  const placed = doc.addImage(2, fixture('red.png'), { width: units.inches(2), align: 'center' })
  t.alike(placed, { index: 2, width: 1828800, height: 914400 })
  const again = reopen(doc)
  t.is(para(blocks(again)[2]).images, 1)
  const files = unzip(again.save())
  t.is(Object.keys(files).filter((n) => n.startsWith('word/media/')).length, 2)
  const xml = decode(files['word/document.xml'])
  t.ok(xml.includes('<wp:extent cx="1828800" cy="914400"/>'))
})

test('addChart and setChartData work on Word documents', (t) => {
  const doc = open()
  doc.addChart(12, {
    type: 'bar',
    title: 'Margin',
    categories: ['2025', '2026'],
    series: [{ name: 'Margin', values: [25, 36] }]
  })
  let read = para(blocks(reopen(doc))[12])
  t.is(read.charts?.[0].type, 'bar')
  doc.setChartData(12, {
    categories: ['2025', '2026', '2027'],
    series: [{ name: 'Margin', values: [25, 36, 40] }]
  })
  doc.setChartTitle(12, 'Margin, forecast')
  read = para(blocks(reopen(doc))[12])
  t.alike(read.charts?.[0].series[0].values, [25, 36, 40])
  t.is(read.charts?.[0].title, 'Margin, forecast')
  const files = unzip(doc.save())
  t.ok(files['word/charts/chart1.xml'])
  t.ok(files['word/embeddings/Microsoft_Excel_Worksheet1.xlsx'])
  t.exception(() => doc.setChartData(3, { categories: ['a'], series: [] }), /holds no chart/)
})

test('create builds a document with the standard styles', (t) => {
  const doc = docx.create({ title: 'Plan', page: 'a4', fonts: { body: 'Georgia' }, size: 12 })
  t.is(doc.blockCount, 1)
  t.alike(doc.paragraphStyles(), [
    'Normal',
    'Title',
    'Subtitle',
    'Heading 1',
    'Heading 2',
    'Heading 3',
    'List Bullet',
    'List Number',
    'Quote'
  ])
  doc.appendParagraphs([
    { text: 'Launch Plan', style: 'Title' },
    { text: 'Goals', style: 'Heading 1' },
    { text: 'Ship the beta', style: 'List Number' },
    { text: 'Collect feedback', style: 'List Number' },
    { runs: ['Owner: ', { text: 'Product', bold: true }] }
  ])
  doc.addTable(doc.blockCount, {
    rows: [
      ['Milestone', 'Date'],
      ['Beta', 'May']
    ]
  })
  doc.addImage(doc.blockCount, fixture('red.png'))
  const again = reopen(doc)
  const read = blocks(again)
  t.alike(
    read.slice(0, 5).map((b) => para(b).style),
    ['Title', 'Heading 1', 'List Number', 'List Number', 'Normal']
  )
  t.ok(para(read[2]).list)
  t.is(para(read[4]).text, 'Owner: Product')
  t.is(read[5].type, 'table')
  t.is(para(read[6]).images, 1)

  const files = unzip(again.save())
  t.ok(files['word/styles.xml'])
  t.ok(files['word/numbering.xml'])
  t.ok(files['word/settings.xml'])
  t.ok(decode(files['word/document.xml']).includes('<w:pgSz w:w="11906" w:h="16838"/>'))
  t.ok(decode(files['word/styles.xml']).includes('<w:rFonts w:ascii="Georgia"'))
  t.ok(decode(files['docProps/core.xml']).includes('<dc:title>Plan</dc:title>'))
  t.exception(() => docx.create({ page: 'legal' } as never), /page must be/)
})

test('an edited document survives repeated open and save', (t) => {
  let doc = open()
  doc.replaceSection(2, { body: ['Short.'] })
  doc.addImage(0, fixture('red.png'))
  doc = reopen(doc)
  doc.deleteBlocks(0)
  doc.setCell({ table: 6, row: 0, column: 0 }, 'Measure')
  doc = reopen(doc)
  t.alike(texts(doc).slice(0, 4), [
    'Quarterly Report',
    'Prepared by the finance team.',
    'Revenue',
    'Short.'
  ])
  t.is(table(blocks(doc)[6]).rows[0][0], 'Measure')
})
