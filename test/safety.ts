import test from 'brittle'
import { pptx, docx, xlsx, DocsError } from '../lib/index.ts'
import * as xml from '../lib/xml.ts'
import type { Document, Element, Node } from '../lib/xml.ts'
import { fixture, decode, unzip, zipFiles } from './helpers/engine.ts'
import { encode } from '../lib/utf8.ts'

const W = xml.ns.w
const A = xml.ns.a

interface Insert {
  markup: string
  before?: Node | null
}

function insertXml(doc: Document, parent: Element, insert: Insert): void {
  const wrapper = xml.root(
    xml.parse(`<root xmlns:w="${W}" xmlns:a="${A}" xmlns:r="${xml.ns.r}">${insert.markup}</root>`)
  )
  for (const node of Array.from(wrapper.childNodes)) {
    parent.insertBefore(doc.importNode(node, true), insert.before ?? null)
  }
}

function bodyParagraph(doc: ReturnType<typeof docx.open>, index: number): Element {
  const body = xml.descendants(doc.package.xml(doc.main), W, 'body')[0]
  return xml.children(body, W, 'p')[index]
}

test('Word refuses to rewrite text with tracked changes', (t) => {
  const doc = docx.open(fixture('report.docx'))
  insertXml(doc.package.xml(doc.main), bodyParagraph(doc, 3), {
    markup: '<w:ins w:id="1" w:author="A"><w:r><w:t> Added.</w:t></w:r></w:ins>'
  })
  doc.package.touch(doc.main)
  const block = doc.read().blocks[3]
  t.is(
    block.type === 'paragraph' ? block.text : null,
    'Revenue grew twelve percent year over year. Added.'
  )
  try {
    doc.setText(3, 'Replaced')
    t.fail('expected TRACKED_CHANGES')
  } catch (err) {
    t.ok(err instanceof DocsError && err.code === 'TRACKED_CHANGES')
    t.ok(err instanceof Error && /accept or reject them in Word first/.test(err.message))
  }
})

test('Word reports links, images and fields that a rewrite drops', (t) => {
  const doc = docx.open(fixture('report.docx'))
  insertXml(doc.package.xml(doc.main), bodyParagraph(doc, 1), {
    markup:
      '<w:hyperlink w:anchor="top"><w:r><w:t>link</w:t></w:r></w:hyperlink><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>'
  })
  doc.package.touch(doc.main)
  t.alike(doc.setText(1, 'Plain now'), { changed: true, removed: { hyperlinks: 1, fields: 1 } })
  t.alike(doc.setText(13, 'Caption'), { changed: true, removed: { images: 1 } })
  t.alike(doc.setText(12, 'Calm.'), { changed: true })
})

test('PowerPoint does not spread a link over new text and reports it', (t) => {
  const deck = pptx.open(fixture('basic.pptx'))
  const slide = deck.package.xml('/ppt/slides/slide1.xml')
  const run = xml.descendants(slide, A, 'r')[0]
  insertXml(slide, run, {
    markup: '<a:rPr b="1"><a:hlinkClick r:id="rId99"/></a:rPr>',
    before: run.firstChild
  })
  deck.package.touch('/ppt/slides/slide1.xml')
  const result = deck.setText({ slide: 0, shape: 'title' }, ['First', 'Second'])
  t.alike(result, { paragraphs: 2, changed: true, removed: { hyperlinks: 1 }, overflow: false })
  const out = decode(deck.package.bytes('/ppt/slides/slide1.xml'))
  t.absent(out.includes('hlinkClick'))
  t.ok(out.includes('<a:rPr b="1"/><a:t>Second</a:t>'))
})

test('Excel checks that = text is a formula', (t) => {
  const book = xlsx.open(fixture('shared.xlsx'))
  t.exception(
    () => book.setCells(0, { A1: '=)' }),
    /"=\)" is not a formula \(a \) without a matching \(\); to store it as text use \{ value: "=\)" \}/
  )
  t.exception(() => book.setCells(0, { A1: '=SUM(B2:B3' }), /a \( without a matching \)/)
  t.exception(() => book.setCells(0, { A1: '=B2+' }), /ends with an operator/)
  t.exception(() => book.setCells(0, { A1: '="open' }), /unclosed string/)
  t.exception(() => book.setCells(0, { A1: { formula: '=' } }), /it is empty/)
  t.execution(() =>
    book.setCells(0, { A1: '=IF(B2>0,"a (b","c")', A2: "='My Sheet'!A1", A3: { value: '=)' } })
  )
  t.is(book.get(0, 'A3').value, '=)')
})

test('Excel dates follow the 1900 leap-year convention', (t) => {
  const book = xlsx.create()
  book.setCells(0, {
    A1: { date: '1900-01-01' },
    A2: { date: '1900-02-28' },
    A3: { date: '1900-03-01' },
    A4: { value: 60, format: 'yyyy-mm-dd' }
  })
  const again = xlsx.open(book.save())
  t.alike(
    again.read(0).rows.map((r) => r[0]),
    ['1900-01-01', '1900-02-28', '1900-03-01', '1900-02-29']
  )
  const sheet = decode(unzip(again.save())['xl/worksheets/sheet1.xml'])
  for (const serial of ['<v>1</v>', '<v>59</v>', '<v>61</v>']) {
    t.ok(sheet.includes(serial), serial)
  }
})

test('saving drops stale summaries and records the modification when asked', (t) => {
  const deck = pptx.open(fixture('basic.pptx'))
  deck.addSlide({ title: 'New' })
  const files = unzip(
    deck.save({ modified: new Date('2026-09-25T12:00:00.123Z'), modifiedBy: 'QVAC' })
  )
  const app = decode(files['docProps/app.xml'])
  t.ok(app.includes('<Slides>4</Slides>'))
  t.absent(app.includes('TitlesOfParts'))
  t.absent(app.includes('HeadingPairs'))
  const core = decode(files['docProps/core.xml'])
  t.ok(
    core.includes(
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-09-25T12:00:00Z</dcterms:modified>'
    )
  )
  t.ok(core.includes('<cp:lastModifiedBy>QVAC</cp:lastModifiedBy>'))
  const book = xlsx.open(fixture('shared.xlsx'))
  book.renameSheet(0, 'Costs')
  t.absent(decode(unzip(book.save())['docProps/app.xml']).includes('<vt:lpstr>Budget</vt:lpstr>'))
  const untouched = xlsx.open(fixture('shared.xlsx'))
  untouched.setCells(0, { B2: 1 })
  t.ok(decode(unzip(untouched.save())['docProps/app.xml']).includes('<vt:lpstr>Budget</vt:lpstr>'))
  const notADate: unknown = 'yesterday'
  t.exception(
    () => docx.create().save({ modified: notADate instanceof Date ? notADate : new Date(NaN) }),
    /modified must be a valid Date/
  )
})

test('characters outside the Basic Multilingual Plane survive an edit on every runtime', (t) => {
  const deck = pptx.open(fixture('basic.pptx'))
  deck.setText({ slide: 0, shape: 'title' }, 'Launch \u{1F680} \u{1D432}')
  const bytes = unzip(deck.save())['ppt/slides/slide1.xml']
  const hex = Array.from(bytes)
    .map((b) => b.toString(16).padStart(2, '0'))
    .join('')
  t.ok(hex.includes('f09f9a80'))
  t.ok(hex.includes('f09d90b2'))
  t.is(
    pptx.open(deck.save()).read()[0].shapes[0].paragraphs?.[0].text,
    'Launch \u{1F680} \u{1D432}'
  )
})

test('a slide nested deep behind comments is refused as unsafe, not with a stack overflow', (t) => {
  const files = unzip(fixture('basic.pptx'))
  const name = 'ppt/slides/slide1.xml'
  const group =
    '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
  const deep = '<!--</p:grpSp>-->' + group
  files[name] = encode(
    decode(files[name]).replace(
      '<p:grpSpPr/>',
      `<p:grpSpPr/>${deep.repeat(50000)}${'</p:grpSp>'.repeat(50000)}`
    )
  )
  let err: unknown = null
  try {
    pptx.open(zipFiles(files)).read()
  } catch (caught) {
    err = caught
  }

  t.ok(err instanceof DocsError)
  t.is(err instanceof DocsError ? err.code : null, 'UNSAFE_XML')
})
