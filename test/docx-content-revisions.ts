import test from 'brittle'
import * as docx from '../lib/docx/index.ts'
import * as utf8 from '../lib/utf8.ts'
import { decode, unzip, zipFiles } from './helpers/engine.ts'
import type { Block, Document, ParagraphBlock } from '../lib/docx/index.ts'

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const BY = 'w:author="Ana" w:date="2026-01-01T00:00:00Z"'
const reopen = (doc: Document): Document => docx.open(doc.save())

const TRACKED = [
  `<w:p><w:r><w:t xml:space="preserve">Keep </w:t></w:r><w:ins w:id="1" ${BY}><w:r><w:t>added</w:t></w:r></w:ins><w:del w:id="2" ${BY}><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>`,
  `<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="3" ${BY}><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr><w:r><w:rPr><w:b/><w:rPrChange w:id="4" ${BY}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>Styled</w:t></w:r></w:p>`,
  `<w:p><w:pPr><w:rPr><w:del w:id="5" ${BY}/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">Joined </w:t></w:r></w:p>`,
  `<w:p><w:r><w:t>tail</w:t></w:r></w:p>`,
  `<w:p><w:pPr><w:rPr><w:moveFrom w:id="8" ${BY}/></w:rPr></w:pPr><w:moveFrom w:id="6" ${BY}><w:r><w:t>Moved</w:t></w:r></w:moveFrom></w:p>`,
  `<w:p><w:r><w:t>Middle</w:t></w:r></w:p>`,
  `<w:p><w:pPr><w:rPr><w:moveTo w:id="9" ${BY}/></w:rPr></w:pPr><w:moveTo w:id="7" ${BY}><w:r><w:t>Moved</w:t></w:r></w:moveTo></w:p>`,
  `<w:p><w:r><w:t>End</w:t></w:r></w:p>`
].join('')

function tracked(): Document {
  const files = unzip(docx.create().save())
  files['word/document.xml'] = utf8.encode(
    `<w:document xmlns:w="${W}"><w:body>${TRACKED}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`
  )

  return docx.open(zipFiles(files))
}

function para(block: Block | undefined): ParagraphBlock {
  if (!block || block.type !== 'paragraph') {
    throw new Error('expected a paragraph block')
  }

  return block
}

const texts = (doc: Document): string[] => doc.read().blocks.map((b) => para(b).text)
const body = (doc: Document): string => decode(unzip(doc.save())['word/document.xml'])

test('acceptChanges keeps insertions, drops deletions and joins deleted paragraph marks', (t) => {
  const doc = tracked()
  t.alike(doc.acceptChanges(), {
    insertions: 1,
    deletions: 1,
    moves: 2,
    formatting: 2,
    paragraphs: 3,
    rows: 0
  })
  const reread = reopen(doc)
  t.alike(texts(reread), ['Keep added', 'Styled', 'Joined tail', 'Middle', 'Moved', 'End'])
  const xml = body(reread)
  t.ok(xml.includes('<w:jc w:val="center"/>'))
  t.ok(xml.includes('<w:b/>'))
  for (const local of ['ins', 'del', 'moveFrom', 'moveTo', 'pPrChange', 'rPrChange', 'delText']) {
    t.absent(xml.includes(`<w:${local} `) || xml.includes(`<w:${local}>`), local)
  }
})

test('rejectChanges restores deletions and the old formatting', (t) => {
  const doc = tracked()
  const counts = doc.rejectChanges()
  t.is(counts.insertions, 1)
  t.is(counts.deletions, 1)
  t.is(counts.formatting, 2)
  t.alike(texts(reopen(doc)), ['Keep gone', 'Styled', 'Joined ', 'tail', 'Moved', 'Middle', 'End'])
  const xml = body(doc)
  t.ok(xml.includes('<w:jc w:val="right"/>'))
  t.ok(xml.includes('<w:i/>'))
  t.absent(xml.includes('<w:b/>'))
  t.ok(xml.includes('<w:t>gone</w:t>'))
})

test('acceptChanges with in touches only that block', (t) => {
  const doc = tracked()
  t.exception(() => doc.setText(0, 'Rewritten'), /tracked changes/)
  const counts = doc.acceptChanges({ in: 0 })
  t.is(counts.insertions + counts.deletions, 2)
  t.is(counts.formatting, 0)
  t.alike(doc.setText(0, 'Rewritten'), { changed: true })
  t.exception(() => doc.setText(1, 'x'), /tracked changes/)
  t.is(para(reopen(doc).read().blocks[0]).text, 'Rewritten')
})

test('accepting a document without changes changes nothing', (t) => {
  const doc = docx.create()
  doc.appendParagraphs(['Plain'])
  const saved = doc.save()
  const again = docx.open(saved)
  t.alike(again.acceptChanges(), {
    insertions: 0,
    deletions: 0,
    moves: 0,
    formatting: 0,
    paragraphs: 0,
    rows: 0
  })
  t.absent(again.changed)
})

test('tracked rows are kept or removed', (t) => {
  const files = unzip(docx.create().save())
  const row = (mark: string, value: string): string =>
    `<w:tr><w:trPr>${mark}</w:trPr><w:tc><w:p><w:r><w:t>${value}</w:t></w:r></w:p></w:tc></w:tr>`
  const tbl = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>${row('', 'a')}${row(`<w:ins w:id="1" ${BY}/>`, 'b')}${row(`<w:del w:id="2" ${BY}/>`, 'c')}</w:tbl><w:p/>`
  files['word/document.xml'] = utf8.encode(
    `<w:document xmlns:w="${W}"><w:body>${tbl}<w:sectPr/></w:body></w:document>`
  )
  const accepted = docx.open(zipFiles(files))
  t.is(accepted.acceptChanges().rows, 2)
  t.alike(accepted.read().blocks[0], { index: 0, type: 'table', rows: [['a'], ['b']] })
  const rejected = docx.open(zipFiles(files))
  rejected.rejectChanges()
  t.alike(rejected.read().blocks[0], { index: 0, type: 'table', rows: [['a'], ['c']] })
})
