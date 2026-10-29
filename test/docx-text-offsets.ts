import test from 'brittle'
import * as docx from '../lib/docx/index.ts'
import * as utf8 from '../lib/utf8.ts'
import { DocsError } from '../lib/errors.ts'
import { fixture, decode, unzip, zipFiles } from './helpers/engine.ts'
import type { Block, Document, ParagraphBlock } from '../lib/docx/index.ts'

const open = (): Document => docx.open(fixture('report.docx'))
const reopen = (doc: Document): Document => docx.open(doc.save())

function para(block: Block | undefined): ParagraphBlock {
  if (!block || block.type !== 'paragraph') {
    throw new Error('expected a paragraph block')
  }

  return block
}

function at(doc: Document, index: number): ParagraphBlock {
  return para(doc.read({ from: index, count: 1, detail: 'runs' }).blocks[0])
}

function failure(fn: () => unknown): DocsError {
  try {
    fn()
  } catch (err) {
    if (err instanceof DocsError) {
      return err
    }

    throw err
  }

  throw new Error('expected a DocsError')
}

function withBody(body: string): Document {
  const files = unzip(docx.create().save())
  const xml = decode(files['word/document.xml']).replace('<w:body><w:p/>', `<w:body>${body}`)
  files['word/document.xml'] = utf8.encode(xml)
  return docx.open(zipFiles(files))
}

function bodyXml(doc: Document): string {
  return decode(unzip(doc.save())['word/document.xml'])
}

const run = (value: string, rPr = ''): string =>
  `<w:r>${rPr}<w:t xml:space="preserve">${value}</w:t></w:r>`

const FIELD =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'

test('insertText types into the middle of a run, keeping it one run', (t) => {
  const doc = open()
  t.alike(doc.insertText(3, { offset: 3, text: 'XX' }), { length: 45, offset: 5 })
  t.alike(at(reopen(doc), 3).runs, [
    { text: 'RevXXenue grew ' },
    { text: 'twelve percent', bold: true },
    { text: ' year over year.' }
  ])
})

test('insertText at a run boundary takes the formatting of the run before', (t) => {
  const doc = open()
  doc.insertText(3, { offset: 13, text: 'A' })
  doc.insertText(3, { offset: 28, text: 'B' })
  t.alike(at(reopen(doc), 3).runs, [
    { text: 'Revenue grew A' },
    { text: 'twelve percentB', bold: true },
    { text: ' year over year.' }
  ])
})

test('insertText at 0 and at the end uses the first and last runs', (t) => {
  const doc = open()
  doc.insertText(3, { offset: 0, text: '> ' })
  const { length } = doc.insertText(3, { offset: 45, text: ' <' })
  t.is(length, 47)
  const block = at(reopen(doc), 3)
  t.is(block.text, '> Revenue grew twelve percent year over year. <')
  t.alike(block.runs, [
    { text: '> Revenue grew ' },
    { text: 'twelve percent', bold: true },
    { text: ' year over year. <' }
  ])
})

test('insertText turns newlines and tabs into breaks and tabs', (t) => {
  const doc = open()
  t.alike(doc.insertText(1, { offset: 8, text: 'a\tb\r\nc' }), { length: 34, offset: 13 })
  const xml = bodyXml(doc)
  t.ok(xml.includes('<w:tab/>'))
  t.ok(xml.includes('<w:br/>'))
  t.is(at(reopen(doc), 1).text, 'Prepareda\tb\nc by the finance team.')
})

test('insertText with format puts the text in its own run over the inherited format', (t) => {
  const doc = open()
  doc.insertText(3, { offset: 3, text: 'B', format: { bold: true } })
  doc.insertText(3, { offset: 20, text: 'n', format: { bold: false, italic: true } })
  doc.insertText(3, { offset: 16, text: 's', format: { bold: true } })
  t.alike(at(reopen(doc), 3).runs, [
    { text: 'Rev' },
    { text: 'B', bold: true },
    { text: 'enue grew ' },
    { text: 'twselve', bold: true },
    { text: 'n', italic: true },
    { text: ' percent', bold: true },
    { text: ' year over year.' }
  ])
})

test('insertText into an empty paragraph takes the paragraph mark formatting', (t) => {
  const doc = withBody('<w:p><w:pPr><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr></w:pPr></w:p>')
  t.alike(doc.runFormat(0, { offset: 0 }), {
    bold: true,
    italic: false,
    underline: false,
    size: 11,
    color: 'FF0000',
    font: 'Calibri'
  })
  doc.insertText(0, { offset: 0, text: 'Hi' })
  t.alike(at(reopen(doc), 0).runs, [{ text: 'Hi', bold: true, color: 'FF0000' }])
})

test('runFormat reports the format typing at an offset would get, styles included', (t) => {
  const doc = open()
  t.is(doc.runFormat(3, { offset: 13 }).bold, false)
  t.is(doc.runFormat(3, { offset: 14 }).bold, true)
  t.is(doc.runFormat(3, { offset: 27 }).bold, true)
  const created = docx.create()
  created.appendParagraphs([{ text: 'Plan', style: 'Heading 1' }])
  const heading = created.runFormat(0, { offset: 2 })
  t.is(heading.size, 16)
  t.is(heading.color, '2F5496')
})

test('deleteText removes a range across runs and drops runs left empty', (t) => {
  const doc = open()
  t.alike(doc.deleteText(3, { start: 8, end: 28 }), { length: 23, offset: 8 })
  const again = reopen(doc)
  t.alike(at(again, 3).runs, [{ text: 'Revenue year over year.' }])
  t.absent(bodyXml(again).includes('<w:b/>'), 'the emptied bold run is gone')
  const partial = open()
  partial.deleteText(3, { start: 10, end: 30 })
  t.alike(at(reopen(partial), 3).runs, [{ text: 'Revenue grar over year.' }])
})

test('deleting nothing succeeds without changing the document', (t) => {
  const doc = open()
  t.alike(doc.deleteText(3, { start: 5, end: 5 }), { length: 43, offset: 5 })
  t.alike(doc.insertText(3, { offset: 5, text: '' }), { length: 43, offset: 5 })
  t.absent(doc.changed)
})

test('splitParagraph then mergeParagraphs gives back the original text and runs', (t) => {
  const doc = open()
  const before = at(doc, 3)
  const split = doc.splitParagraph(3, { offset: 20 })
  t.is(split.first.block, 3)
  t.is(split.second.block, 4)
  t.is(at(doc, 3).text, 'Revenue grew twelve ')
  t.is(at(doc, 4).text, 'percent year over year.')
  t.is(at(doc, 4).style, 'Normal')
  t.alike(doc.mergeParagraphs(3), { block: 3, id: split.first.id, offset: 20, length: 43 })
  const again = reopen(doc)
  t.is(again.blockCount, 14)
  t.alike(at(again, 3).runs, before.runs)
  t.is(at(again, 4).text, 'North led the growth')
})

test('splitParagraph keeps the original id and gives the new paragraph a fresh one', (t) => {
  const doc = docx.create()
  const [id] = doc.appendParagraphs(['Alpha beta']).ids
  const split = doc.splitParagraph({ id }, { offset: 5 })
  t.is(split.first.id, id)
  t.ok(/^[0-9A-F]{8}$/.test(split.second.id))
  t.not(split.second.id, id)
  const again = reopen(doc)
  t.alike(
    again.read().blocks.map((b) => b.id),
    [id, split.second.id]
  )
  t.is(at(again, 1).text, ' beta')
  const old = open()
  t.absent(at(old, 1).id, 'the fixture paragraph has no id yet')
  const stamped = old.splitParagraph(1, { offset: 0 })
  t.is(at(old, 1).id, stamped.first.id)
  t.is(at(old, 1).text, '')
  t.is(at(old, 2).text, 'Prepared by the finance team.')
})

test('splitting at the end of a heading starts the next style', (t) => {
  const doc = docx.create()
  doc.appendParagraphs([
    { text: 'Plan', style: 'Heading 1' },
    { text: 'Item', style: 'List Bullet' }
  ])
  doc.splitParagraph(0, { offset: 4 })
  t.is(at(doc, 1).style, 'Normal')
  doc.splitParagraph(0, { offset: 2 })
  t.is(at(doc, 0).style, 'Heading 1')
  t.is(at(doc, 1).style, 'Heading 1')
  t.is(at(doc, 1).text, 'an')
  doc.splitParagraph(3, { offset: 4 })
  t.is(at(doc, 4).style, 'List Bullet', 'a style without w:next keeps itself')
})

test('formatRange formats exactly the range across three runs', (t) => {
  const doc = open()
  t.alike(doc.formatRange(3, { start: 8, end: 31, format: { italic: true } }), { changed: true })
  t.alike(at(reopen(doc), 3).runs, [
    { text: 'Revenue ' },
    { text: 'grew ', italic: true },
    { text: 'twelve percent', bold: true, italic: true },
    { text: ' yea', italic: true },
    { text: 'r over year.' }
  ])
  doc.formatRange(3, { start: 13, end: 27, format: { bold: false } })
  t.absent(doc.runFormat(3, { offset: 20 }).bold)
  t.alike(doc.formatRange(3, { start: 4, end: 4, format: { bold: true } }), { changed: false })
})

test('offsets outside the text or inside a character are refused with details', (t) => {
  const doc = open()
  const past = failure(() => doc.insertText(3, { offset: 44, text: 'x' }))
  t.is(past.code, 'INVALID_OFFSET')
  t.alike(past.details, { field: 'offset', expected: '0 to 43', got: 44, block: 3, length: 43 })
  t.is(failure(() => doc.insertText(3, { offset: -1, text: 'x' })).code, 'INVALID_OFFSET')
  t.is(failure(() => doc.insertText(3, { offset: 1.5, text: 'x' })).code, 'INVALID_OFFSET')
  const backwards = failure(() => doc.deleteText(3, { start: 9, end: 2 }))
  t.is(backwards.code, 'INVALID_OFFSET')
  t.is(backwards.details?.field, 'end')
  t.is(
    failure(() => doc.formatRange(3, { start: 0, end: 99, format: { bold: true } })).code,
    'INVALID_OFFSET'
  )
  const emoji = withBody(`<w:p>${run('a😀b')}</w:p>`)
  const pair = failure(() => emoji.insertText(0, { offset: 2, text: 'x' }))
  t.is(pair.code, 'INVALID_OFFSET')
  t.alike(pair.details?.candidates, [1, 3])
  t.is(failure(() => emoji.splitParagraph(0, { offset: 2 })).code, 'INVALID_OFFSET')
  t.alike(emoji.insertText(0, { offset: 3, text: '!' }), { length: 5, offset: 4 })
  const untyped = (opts: unknown): unknown => Reflect.apply(doc.insertText, doc, [3, opts])
  t.is(failure(() => untyped({ offset: 0, text: 5 })).code, 'INVALID_TEXT')
  t.is(
    failure(() => untyped({ offset: 0, text: 'x', format: { bold: 'yes' } })).code,
    'INVALID_STYLE'
  )
  t.is(failure(() => untyped({ text: 'x' })).details?.field, 'offset')
})

test('mergeParagraphs refuses a table or a missing next paragraph', (t) => {
  const doc = open()
  const table = failure(() => doc.mergeParagraphs(9))
  t.is(table.code, 'NOT_PARAGRAPH')
  t.is(table.details?.next, 'tbl')
  t.is(failure(() => doc.mergeParagraphs(13)).code, 'BLOCK_NOT_FOUND')
  t.is(failure(() => doc.insertText(10, { offset: 0, text: 'x' })).code, 'NOT_PARAGRAPH')
  t.absent(doc.changed)
})

test('text in a hyperlink stays in it, and typing at its end goes after it', (t) => {
  const doc = withBody(
    `<w:p>${run('See ')}<w:hyperlink w:anchor="top">${run('docs', '<w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>')}</w:hyperlink></w:p>`
  )
  t.is(doc.runFormat(0, { offset: 6 }).link, '#top')
  doc.insertText(0, { offset: 6, text: 'X' })
  t.absent(doc.runFormat(0, { offset: 9 }).link)
  doc.insertText(0, { offset: 9, text: ' now' })
  doc.deleteText(0, { start: 2, end: 5 })
  t.alike(at(reopen(doc), 0).runs, [
    { text: 'Se' },
    { text: 'oXcs', link: '#top' },
    { text: ' now' }
  ])
  doc.deleteText(0, { start: 2, end: 6 })
  t.is(at(doc, 0).text, 'Se now')
  t.absent(bodyXml(doc).includes('<w:hyperlink'), 'an emptied hyperlink is removed')
})

test('splitting inside a hyperlink splits the link into both paragraphs', (t) => {
  const doc = withBody(
    `<w:p>${run('See ')}<w:hyperlink w:anchor="top">${run('the docs')}</w:hyperlink></w:p>`
  )
  doc.splitParagraph(0, { offset: 7 })
  const again = reopen(doc)
  t.alike(at(again, 0).runs, [{ text: 'See ' }, { text: 'the', link: '#top' }])
  t.alike(at(again, 1).runs, [{ text: ' docs', link: '#top' }])
  again.mergeParagraphs(0)
  t.is(at(again, 0).text, 'See the docs')
})

test('fields are kept whole: typing goes around them and their result is not edited', (t) => {
  const doc = withBody(
    `<w:p>${run('Page ')}${FIELD}${run('3')}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`
  )
  t.is(at(doc, 0).text, 'Page 3')
  t.is(failure(() => doc.deleteText(0, { start: 4, end: 6 })).code, 'FIELD_TEXT')
  doc.insertText(0, { offset: 6, text: ' of 9' })
  doc.insertText(0, { offset: 5, text: '#' })
  const xml = bodyXml(doc)
  t.ok(xml.indexOf('of 9') > xml.indexOf('w:fldCharType="end"'), 'typed after the field end')
  t.is(at(reopen(doc), 0).text, 'Page #3 of 9')
  const split = withBody(
    `<w:p>${run('Page ')}${FIELD}${run('3')}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`
  )
  split.splitParagraph(0, { offset: 5 })
  t.is(at(split, 0).text, 'Page ')
  t.is(at(split, 1).text, '3')
  t.ok(bodyXml(split).split('<w:p')[2].includes('w:fldCharType="begin"'), 'the whole field moved')
  const simple = withBody(`<w:p><w:fldSimple w:instr="PAGE">${run('7')}</w:fldSimple></w:p>`)
  simple.insertText(0, { offset: 1, text: '!' })
  t.ok(bodyXml(simple).includes('</w:fldSimple><w:r>'), 'typed after the simple field')
})

test('tracked changes refuse edits that touch them and deleted text has no offsets', (t) => {
  const doc = withBody(
    `<w:p>${run('keep')}<w:del w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:tab/><w:delText>gone</w:delText></w:r></w:del>${run(' more')}<w:ins w:id="3" w:author="A" w:date="2026-01-01T00:00:00Z">${run(' new')}</w:ins></w:p>`
  )
  t.is(at(doc, 0).text, 'keep more new')
  doc.insertText(0, { offset: 6, text: 'X' })
  t.is(at(doc, 0).text, 'keep mXore new')
  t.is(failure(() => doc.deleteText(0, { start: 2, end: 6 })).code, 'TRACKED_CHANGES')
  t.is(failure(() => doc.insertText(0, { offset: 12, text: 'x' })).code, 'TRACKED_CHANGES')
  t.is(
    failure(() => doc.formatRange(0, { start: 9, end: 12, format: { bold: true } })).code,
    'TRACKED_CHANGES'
  )
  doc.formatRange(0, { start: 5, end: 8, format: { bold: true } })
  t.ok(bodyXml(reopen(doc)).includes('<w:delText>gone</w:delText>'))
})

test('bookmarks and comment ranges survive edits, splits and merges', (t) => {
  const doc = withBody(
    `<w:p><w:bookmarkStart w:id="0" w:name="mark"/>${run('Hello')}<w:bookmarkEnd w:id="0"/><w:commentRangeStart w:id="1"/>${run(' world')}<w:commentRangeEnd w:id="1"/></w:p>`
  )
  doc.insertText(0, { offset: 5, text: ',' })
  doc.deleteText(0, { start: 0, end: 1 })
  doc.splitParagraph(0, { offset: 6 })
  doc.mergeParagraphs(0)
  const xml = bodyXml(doc)
  for (const tag of ['bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd']) {
    t.ok(xml.includes(`<w:${tag}`), tag)
  }

  t.is(at(reopen(doc), 0).text, 'ello, world')
})

test('line breaks, tabs and no-break hyphens count one character each', (t) => {
  const doc = withBody(
    `<w:p>${run('a')}<w:r><w:tab/><w:t>b</w:t><w:br/><w:br w:type="page"/><w:noBreakHyphen/><w:t>c</w:t></w:r></w:p>`
  )
  t.is(at(doc, 0).text, 'a\tb\n-c')
  doc.insertText(0, { offset: 5, text: 'X' })
  doc.deleteText(0, { start: 1, end: 2 })
  t.is(at(reopen(doc), 0).text, 'ab\n-Xc')
})

test('cell paragraphs are addressed by cell or by id', (t) => {
  const doc = open()
  const cell = { table: 10, row: 1, column: 0 }
  doc.insertText(cell, { offset: 6, text: ' rate' })
  const split = doc.splitParagraph(cell, { offset: 6 })
  t.is(split.first.block, 10)
  t.is(split.second.block, 10)
  t.alike(doc.read({ from: 10, count: 1 }).blocks[0], {
    index: 10,
    type: 'table',
    rows: [
      ['Metric', '2025', '2026'],
      ['Margin\n rate', '25%', '36%']
    ]
  })
  t.is(doc.runFormat({ id: split.second.id }, { offset: 0 }).bold, false)
  t.is(doc.insertText({ ...cell, paragraph: 1 }, { offset: 5, text: 's' }).length, 6)
  t.is(failure(() => doc.mergeParagraphs({ id: split.second.id })).code, 'BLOCK_NOT_FOUND')
  t.alike(doc.mergeParagraphs({ id: split.first.id }).offset, 6)
  const missing = failure(() => doc.insertText({ ...cell, paragraph: 3 }, { offset: 0, text: 'x' }))
  t.is(missing.code, 'PARAGRAPH_NOT_FOUND')
  t.alike(reopen(doc).read({ from: 10, count: 1 }).blocks[0], {
    index: 10,
    type: 'table',
    rows: [
      ['Metric', '2025', '2026'],
      ['Margin rates', '25%', '36%']
    ]
  })
})
