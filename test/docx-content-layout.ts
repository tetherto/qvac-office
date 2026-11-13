import test from 'brittle'
import * as docx from '../lib/docx/index.ts'
import * as units from '../lib/units.ts'
import * as utf8 from '../lib/utf8.ts'
import { fixture, decode, unzip, zipFiles } from './helpers/engine.ts'
import type { Block, Document, ParagraphBlock } from '../lib/docx/index.ts'

const reopen = (doc: Document): Document => docx.open(doc.save())

function para(block: Block | undefined): ParagraphBlock {
  if (!block || block.type !== 'paragraph') {
    throw new Error('expected a paragraph block')
  }

  return block
}

test('page() reads size, orientation and margins in EMU', (t) => {
  const page = docx.create({ page: 'a4' }).page()
  t.is(page.name, 'a4')
  t.is(page.orientation, 'portrait')
  t.alike(page.size, { width: 11906 * 635, height: 16838 * 635 })
  t.is(page.margins.top, units.inches(1))
  t.is(page.margins.header, units.inches(0.5))
  t.is(page.sections, 1)
})

test('setPage changes size, orientation and margins with units', (t) => {
  const doc = docx.create()
  const result = doc.setPage({
    size: 'a4',
    orientation: 'landscape',
    margins: { top: '0.5in', left: '2cm', footer: 0 }
  })
  t.is(result.sections, 1)
  const page = reopen(doc).page()
  t.is(page.name, 'a4')
  t.is(page.orientation, 'landscape')
  t.alike(page.size, { width: 16838 * 635, height: 11906 * 635 })
  t.is(page.margins.top, units.inches(0.5))
  t.is(page.margins.left, Math.round(units.cm(2) / 635) * 635)
  t.is(page.margins.footer, 0)
  t.is(page.margins.right, units.inches(1))
  const body = decode(unzip(doc.save())['word/document.xml'])
  t.ok(body.includes('w:orient="landscape"'))
})

test('setPage takes a custom size and flips orientation on its own', (t) => {
  const doc = docx.create()
  doc.setPage({ size: { width: '5in', height: '8in' } })
  t.alike(doc.page().size, { width: units.inches(5), height: units.inches(8) })
  t.is(doc.page().name, null)
  doc.setPage({ orientation: 'landscape' })
  t.alike(doc.page().size, { width: units.inches(8), height: units.inches(5) })
  doc.setPage({ orientation: 'portrait' })
  t.is(doc.page().orientation, 'portrait')
})

test('setPage with all changes every section', (t) => {
  const doc = docx.create()
  doc.appendParagraphs(['One', 'Two'])
  const files = unzip(doc.save())
  const xml = decode(files['word/document.xml']).replace(
    /<w:p( [^>]*)?><w:r><w:t>One<\/w:t><\/w:r><\/w:p>/,
    '<w:p$1><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>One</w:t></w:r></w:p>'
  )
  t.ok(xml.includes('<w:sectPr><w:pgSz'))
  const twoSections = docx.open(rewrite(files, xml))
  t.alike(twoSections.setPage({ orientation: 'landscape' }).sections, 1)
  t.alike(twoSections.setPage({ orientation: 'landscape', all: true }).sections, 2)
  const body = decode(unzip(twoSections.save())['word/document.xml'])
  t.is(body.split('w:orient="landscape"').length - 1, 2)
})

test('setPage refuses bad sizes, orientations and margins', (t) => {
  const doc = docx.create()
  t.exception(() => doc.setPage({ size: JSON.parse('"legal"') }), /size must be/)
  t.exception(() => doc.setPage({ orientation: JSON.parse('"sideways"') }), /orientation must be/)
  t.exception(() => doc.setPage({ margins: JSON.parse('{"inside":1}') }), /unknown margin/)
  t.exception(() => doc.setPage({ margins: { top: '1 furlong' } }), /top must be EMU/)
})

test('addTableOfContents inserts a TOC field listing headings and asks Word to update', (t) => {
  const doc = docx.open(fixture('report.docx'))
  const result = doc.addTableOfContents(1, { levels: 2, title: 'Contents' })
  t.alike({ ...result, ids: result.ids.length }, { index: 1, count: 5, levels: 2, ids: 5 })
  const read = reopen(doc).read().blocks
  t.alike(
    read.slice(1, 6).map((b) => para(b).text),
    ['Contents', 'Revenue', 'Regions', 'Costs', 'Outlook']
  )
  t.absent(para(read[1]).heading)
  const files = unzip(doc.save())
  const body = decode(files['word/document.xml'])
  t.ok(body.includes('TOC \\o "1-2" \\h \\z \\u'))
  t.ok(body.includes('w:fldCharType="begin" w:dirty="true"'))
  t.ok(body.includes('w:fldCharType="end"'))
  t.ok(decode(files['word/settings.xml']).includes('<w:updateFields w:val="true"/>'))
})

test('addTableOfContents writes a placeholder when there are no headings', (t) => {
  const doc = docx.create()
  doc.appendParagraphs(['Plain'])
  const placeholder = doc.addTableOfContents(0)
  t.alike(
    { ...placeholder, ids: placeholder.ids.length },
    { index: 0, count: 1, levels: 3, ids: 1 }
  )
  t.ok(reopen(doc).read().blocks[0])
  t.ok(para(doc.read().blocks[0]).text.startsWith('Update the table of contents'))
  t.exception(() => doc.addTableOfContents(0, { levels: 10 }), /levels must be/)
})

function rewrite(files: Record<string, Uint8Array>, body: string): Uint8Array {
  return zipFiles({ ...files, 'word/document.xml': utf8.encode(body) })
}
