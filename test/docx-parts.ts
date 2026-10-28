import test from 'brittle'
import * as docx from '../lib/docx/index.ts'
import * as xml from '../lib/xml.ts'
import * as utf8 from '../lib/utf8.ts'
import { fixture, decode, unzip, zipFiles } from './helpers/engine.ts'
import type { Block, Document, ParagraphBlock } from '../lib/docx/index.ts'

const W = xml.ns.w
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
const reopen = (doc: Document): Document => docx.open(doc.save())
const NAMESPACES = [
  `xmlns:w="${W}"`,
  `xmlns:r="${xml.ns.r}"`,
  `xmlns:wp="${xml.ns.wp}"`,
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:v="urn:schemas-microsoft-com:vml"'
].join(' ')
const HEADER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'

type Files = Record<string, Uint8Array>

function fromBody(body: string, edit?: (files: Files) => void): Document {
  const files = unzip(docx.create().save())
  files['word/document.xml'] = utf8.encode(
    `<w:document ${NAMESPACES}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`
  )
  edit?.(files)

  return docx.open(zipFiles(files))
}

interface ExtraPart {
  name: string
  type: string
  rel: string
  content: string
}

function addPart(files: Files, part: ExtraPart): void {
  files[`word/${part.name}`] = utf8.encode(part.content)
  const rels = decode(files['word/_rels/document.xml.rels'])
  files['word/_rels/document.xml.rels'] = utf8.encode(
    rels.replace('</Relationships>', `${part.rel}</Relationships>`)
  )
  const types = decode(files['[Content_Types].xml'])
  files['[Content_Types].xml'] = utf8.encode(
    types.replace(
      '</Types>',
      `<Override PartName="/word/${part.name}" ContentType="${part.type}"/></Types>`
    )
  )
}

function mainXml(doc: Document): string {
  return decode(unzip(doc.save())['word/document.xml'])
}

function cell(text: string, props = ''): string {
  return `<w:tc>${props}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
}

function para(block: Block | undefined): ParagraphBlock {
  if (!block || block.type !== 'paragraph') {
    throw new Error('expected a paragraph block')
  }

  return block
}

test('list paragraphs take and report a level', (t) => {
  const doc = docx.open(fixture('report.docx'))
  doc.insertParagraphs(6, [
    { text: 'East, in detail', style: 'List Bullet', level: 1 },
    { text: 'Deeper still', style: 'List Bullet', level: 2 }
  ])
  const read = reopen(doc).read().blocks
  t.alike(
    read.slice(4, 8).map((b) => [para(b).text, para(b).level]),
    [
      ['North led the growth', 0],
      ['South held flat', 0],
      ['East, in detail', 1],
      ['Deeper still', 2]
    ]
  )
  const body = decode(unzip(doc.save())['word/document.xml'])
  t.ok(body.includes('<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>'))
  t.exception(
    () => doc.insertParagraphs(1, { text: 'x', level: 1 }),
    /level applies to list paragraphs/
  )
  t.exception(
    () => doc.insertParagraphs(1, { text: 'x', style: 'List Bullet', level: 9 }),
    /from 0 to 8/
  )
})

test('setHeader rewrites the existing header and setFooter creates one', (t) => {
  const doc = docx.open(fixture('report.docx'))
  t.alike(doc.setHeader('Internal only'), { changed: true, parts: 1 })
  t.alike(doc.setFooter(['Finance', 'Page footer']), { changed: true, parts: 1 })
  const again = reopen(doc)
  const read = again.read()
  t.alike(read.headers, ['Internal only'])
  t.alike(read.footers, ['Finance\nPage footer'])
  const files = unzip(again.save())
  t.ok(files['word/footer1.xml'])
  t.ok(decode(files['word/document.xml']).includes('<w:footerReference w:type="default" r:id='))
  t.ok(decode(files['[Content_Types].xml']).includes('/word/footer1.xml'))
})

test('a first-page header turns on the title page setting, and null removes it', (t) => {
  const doc = docx.create()
  doc.appendParagraphs('Body')
  doc.setHeader('Cover header', { type: 'first' })
  let body = decode(unzip(doc.save())['word/document.xml'])
  t.ok(body.includes('<w:headerReference w:type="first"'))
  t.ok(body.includes('<w:titlePg/>'))
  t.ok(body.indexOf('headerReference') < body.indexOf('<w:pgSz'))
  doc.setHeader(null, { type: 'first' })
  const files = unzip(doc.save())
  body = decode(files['word/document.xml'])
  t.absent(body.includes('headerReference'))
  t.absent(files['word/header1.xml'])
  t.exception(
    () => doc.setHeader('x', { type: 'even' } as never),
    /type must be "default" or "first"/
  )
})

test('read returns footnotes and comments with their anchors', (t) => {
  const doc = docx.open(fixture('report.docx'))
  const pkg = doc.package
  pkg.add(
    '/word/footnotes.xml',
    utf8.encode(
      `<w:footnotes xmlns:w="${W}"><w:footnote w:type="separator" w:id="-1"><w:p/></w:footnote><w:footnote w:id="1"><w:p><w:r><w:t>Audited figures.</w:t></w:r></w:p></w:footnote></w:footnotes>`
    ),
    'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'
  )
  pkg.rels(doc.main).add(REL + 'footnotes', '/word/footnotes.xml')
  pkg.add(
    '/word/comments.xml',
    utf8.encode(
      `<w:comments xmlns:w="${W}"><w:comment w:id="7" w:author="Ana"><w:p><w:r><w:t>Check this number.</w:t></w:r></w:p></w:comment></w:comments>`
    ),
    'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
  )
  pkg.rels(doc.main).add(REL + 'comments', '/word/comments.xml')
  const body = xml.descendants(pkg.xml(doc.main), W, 'body')[0]
  const p = xml.children(body, W, 'p')[3]
  const extra = xml.root(
    xml.parse(
      `<x xmlns:w="${W}"><w:commentRangeStart w:id="7"/><w:r><w:footnoteReference w:id="1"/></w:r><w:commentRangeEnd w:id="7"/></x>`
    )
  )
  for (const node of xml.children(extra)) {
    p.appendChild(pkg.xml(doc.main).importNode(node, true))
  }

  pkg.touch(doc.main)

  const read = reopen(doc).read()
  t.alike(read.footnotes, [{ id: '1', text: 'Audited figures.' }])
  t.alike(read.comments, [{ id: '7', author: 'Ana', text: 'Check this number.' }])
  t.alike(para(read.blocks[3]).footnotes, ['1'])
  t.alike(para(read.blocks[3]).comments, ['7'])
  t.alike(docx.open(fixture('report.docx')).read().comments, [])
})

test('setText turns line and vertical tabs into breaks and drops characters XML cannot hold', (t) => {
  const doc = docx.create()
  doc.setText(0, 'a\u000bb\u0001c\td\r\ne\u0000')
  doc.appendParagraphs('x\u001fy')
  const body = mainXml(doc)
  t.absent(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(body))
  const read = reopen(doc).read().blocks
  t.is(para(read[0]).text, 'a\nbc\td\ne')
  t.is(para(read[1]).text, 'xy')
})

test('setCell addresses cells the way read reports them', (t) => {
  const doc = fromBody(
    '<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid>' +
      `<w:tr>${cell('Merged', '<w:tcPr><w:gridSpan w:val="2"/></w:tcPr>')}${cell('C')}</w:tr>` +
      `<w:tr>${cell('A')}${cell('B')}${cell('C2')}</w:tr></w:tbl><w:p/>`
  )
  const table = (d: Document): string[][] => {
    const block = d.read().blocks[0]
    return block.type === 'table' ? block.rows : []
  }
  t.alike(table(doc)[0], ['Merged', 'C'])
  doc.setCell({ table: 0, row: 0, column: 1 }, 'X')
  t.alike(table(reopen(doc))[0], ['Merged', 'X'])
  t.exception(() => doc.setCell({ table: 0, row: 0, column: 2 }, 'Y'), /the row has 2/)
})

test('cells inside content controls are read and replaced, and added rows get no duplicate ids', (t) => {
  const sdt = (id: number, inner: string): string =>
    `<w:sdt><w:sdtPr><w:id w:val="${id}"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`
  const doc = fromBody(
    '<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>' +
      `<w:tr>${cell('Name')}${cell('Score')}</w:tr>` +
      `<w:tr>${cell('Ann')}<w:tc>${sdt(11, '<w:p><w:r><w:t>3</w:t></w:r></w:p>')}</w:tc></w:tr>` +
      sdt(12, `<w:tr>${cell('Bob')}${sdt(13, cell('4'))}</w:tr>`) +
      '</w:tbl><w:p/>'
  )
  const rows = (d: Document): string[][] => {
    const block = d.read().blocks[0]
    return block.type === 'table' ? block.rows : []
  }
  t.alike(rows(doc), [
    ['Name', 'Score'],
    ['Ann', '3'],
    ['Bob', '4']
  ])
  doc.setCell({ table: 0, row: 1, column: 1 }, '7')
  doc.setCell({ table: 0, row: 2, column: 1 }, '8')
  doc.addRow(0, ['Cy', '9'], { at: 2 })
  doc.addRow(0, ['Di', '5'])
  const again = reopen(doc)
  t.alike(rows(again), [
    ['Name', 'Score'],
    ['Ann', '7'],
    ['Cy', '9'],
    ['Bob', '8'],
    ['Di', '5']
  ])
  const body = mainXml(again)
  const ids = [...body.matchAll(/<w:id w:val="(\d+)"\/>/g)].map((m) => m[1])
  t.is(new Set(ids).size, ids.length)
  t.absent(body.includes('showingPlcHdr'))
  again.deleteRow({ table: 0, row: 4 })
  again.deleteRow({ table: 0, row: 3 })
  t.alike(
    rows(reopen(again)).map((r) => r[0]),
    ['Name', 'Ann', 'Cy']
  )
  t.absent(mainXml(again).includes('w:val="12"'))
})

test('addRow does not copy tracked changes from the row it models', (t) => {
  const doc = fromBody(
    '<w:tbl><w:tblGrid><w:gridCol/></w:tblGrid>' +
      `<w:tr>${cell('Kept')}</w:tr>` +
      '<w:tr><w:trPr><w:del w:id="1" w:author="Ann" w:date="2026-01-01T00:00:00Z"/></w:trPr><w:tc><w:p><w:del w:id="2" w:author="Ann"><w:r><w:delText>Gone</w:delText></w:r></w:del><w:ins w:id="3" w:author="Ann"><w:r><w:t>New</w:t></w:r></w:ins></w:p></w:tc></w:tr>' +
      '</w:tbl><w:p/>'
  )
  doc.addRow(0, ['Added'])
  doc.setCell({ table: 0, row: 2, column: 0 }, 'Edited')
  const body = mainXml(doc)
  const added = body.slice(body.lastIndexOf('<w:tr>'))
  t.absent(added.includes('w:del'))
  t.absent(added.includes('w:ins'))
  t.ok(added.includes('Edited'))
  t.is((body.match(/w:author="Ann"/g) ?? []).length, 3)
})

test('setText reports content controls, equations, objects and breaks it drops, and counts an image once', (t) => {
  const doc = fromBody(
    '<w:p><w:r><w:t>Before</w:t></w:r>' +
      '<w:sdt><w:sdtPr><w:id w:val="4"/></w:sdtPr><w:sdtContent><w:r><w:t>control</w:t></w:r></w:sdtContent></w:sdt>' +
      '<m:oMathPara><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara>' +
      '<m:oMath><m:r><m:t>y</m:t></m:r></m:oMath>' +
      '<w:r><w:object/></w:r>' +
      '<w:r><w:br w:type="page"/><w:lastRenderedPageBreak/><w:br/></w:r>' +
      '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing/></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>' +
      '</w:p>'
  )
  t.alike(doc.setText(0, 'After'), {
    changed: true,
    removed: { images: 1, contentControls: 1, equations: 2, objects: 1, breaks: 1 }
  })
})

test('setText does not carry hyperlink styling onto the new text', (t) => {
  const link =
    '<w:hyperlink w:anchor="top"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:b/><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>link</w:t></w:r></w:hyperlink>'
  const doc = fromBody(
    `<w:p>${link}</w:p><w:p>${link}<w:r><w:rPr><w:i/></w:rPr><w:t> tail</w:t></w:r></w:p>`
  )
  t.alike(doc.setText(0, 'plain'), { changed: true, removed: { hyperlinks: 1 } })
  t.alike(doc.setText(1, 'plain'), { changed: true, removed: { hyperlinks: 1 } })
  const paragraphs = mainXml(doc).split('<w:p>').slice(1)
  t.is(paragraphs[0].split('</w:p>')[0], '<w:r><w:rPr><w:b/></w:rPr><w:t>plain</w:t></w:r>')
  t.is(paragraphs[1].split('</w:p>')[0], '<w:r><w:rPr><w:i/></w:rPr><w:t>plain</w:t></w:r>')
})

test('save keeps parts nothing references until an edit orphans one', (t) => {
  const header = {
    name: 'header9.xml',
    type: HEADER_TYPE,
    rel: `<Relationship Id="rId90" Type="${REL}header" Target="header9.xml"/>`,
    content: `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Unused</w:t></w:r></w:p></w:hdr>`
  }
  const doc = fromBody('<w:p><w:r><w:t>Body</w:t></w:r></w:p>', (files) => addPart(files, header))
  let files = unzip(doc.save())
  t.ok(files['word/header9.xml'])
  t.ok(decode(files['word/_rels/document.xml.rels']).includes('rId90'))
  const image = doc.addImage(1, fixture('red.png'))
  files = unzip(doc.save())
  t.ok(Object.keys(files).some((name) => name.startsWith('word/media/')))
  doc.deleteBlocks(image.index)
  files = unzip(doc.save())
  t.absent(Object.keys(files).some((name) => name.startsWith('word/media/')))
  t.ok(files['word/header9.xml'])
  t.ok(decode(files['word/_rels/document.xml.rels']).includes('rId90'))
})

test('an edit that orphans a part leaves parts nothing referenced at open', (t) => {
  const doc = fromBody('<w:p><w:r><w:t>Body</w:t></w:r></w:p>', (files) => {
    files['word/stray.xml'] = utf8.encode(`<w:hdr xmlns:w="${W}"/>`)
  })
  const image = doc.addImage(1, fixture('red.png'))
  doc.save()
  doc.deleteBlocks(image.index)
  const files = unzip(doc.save())
  t.absent(Object.keys(files).some((name) => name.startsWith('word/media/')))
  t.ok(files['word/stray.xml'])
})

test('new drawing ids do not clash with ids in headers', (t) => {
  const header = {
    name: 'header1.xml',
    type: HEADER_TYPE,
    rel: `<Relationship Id="rId90" Type="${REL}header" Target="header1.xml"/>`,
    content: `<w:hdr xmlns:w="${W}" xmlns:wp="${xml.ns.wp}"><w:p><w:r><w:drawing><wp:inline><wp:docPr id="41" name="Logo"/></wp:inline></w:drawing></w:r></w:p></w:hdr>`
  }
  const doc = fromBody('<w:p/>', (files) => addPart(files, header))
  doc.addImage(0, fixture('red.png'))
  t.ok(/<wp:docPr id="42"/.test(mainXml(doc)))
})

test('numId 0 turns numbering off, including for a list style', (t) => {
  const doc = fromBody(
    '<w:p><w:pPr><w:pStyle w:val="ListBullet"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>Not a list</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>Plain</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:t>Bullet</w:t></w:r></w:p>'
  )
  const blocks = doc.read().blocks.map(para)
  t.alike(
    blocks.map((b) => b.list),
    [undefined, undefined, true]
  )
})

test('created documents name QVAC as the application', (t) => {
  const app = decode(unzip(docx.create().save())['docProps/app.xml'])
  t.ok(app.includes('<Application>QVAC</Application>'))
  t.absent(app.includes('bare-docs'))
})
