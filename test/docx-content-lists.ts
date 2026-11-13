import test from 'brittle'
import * as docx from '../lib/docx/index.ts'
import * as utf8 from '../lib/utf8.ts'
import { decode, unzip, zipFiles } from './helpers/engine.ts'
import type { Block, Document, ParagraphBlock } from '../lib/docx/index.ts'

const reopen = (doc: Document): Document => docx.open(doc.save())
const numberingXml = (doc: Document): string => decode(unzip(doc.save())['word/numbering.xml'])

function para(block: Block | undefined): ParagraphBlock {
  if (!block || block.type !== 'paragraph') {
    throw new Error('expected a paragraph block')
  }

  return block
}

function numIds(doc: Document): string[] {
  const body = decode(unzip(doc.save())['word/document.xml'])
  return [...body.matchAll(/<w:numId w:val="(\d+)"\/>/g)].map((m) => m[1])
}

test('explicit bullet and numbered lists report as lists with levels', (t) => {
  const doc = docx.create()
  doc.appendParagraphs([
    { text: 'Apples', list: 'bullet' },
    { text: 'Green', list: 'bullet', level: 1 },
    { text: 'First', list: 'number' },
    { text: 'Second', list: 'number' }
  ])
  const read = reopen(doc).read().blocks.map(para)
  t.alike(
    read.map((b) => [b.text, b.list, b.level]),
    [
      ['Apples', true, 0],
      ['Green', true, 1],
      ['First', true, 0],
      ['Second', true, 0]
    ]
  )
})

test('numbered paragraphs share an instance within a call and restart in a new call', (t) => {
  const doc = docx.create()
  doc.appendParagraphs([
    { text: 'One', list: 'number' },
    { text: 'Two', list: 'number' }
  ])
  doc.appendParagraphs([{ text: 'Again', list: 'number' }])
  const ids = numIds(doc)
  t.is(ids[0], ids[1])
  t.not(ids[2], ids[0])
  t.ok(numberingXml(doc).includes('<w:startOverride w:val="1"/>'))
})

test('a plain paragraph between numbered ones starts a new instance', (t) => {
  const doc = docx.create()
  doc.appendParagraphs([{ text: 'One', list: 'number' }, 'Break', { text: 'Two', list: 'number' }])
  const ids = numIds(doc)
  t.not(ids[0], ids[1])
})

test('bullets reuse one numbering instance across calls', (t) => {
  const doc = docx.create()
  doc.appendParagraphs([{ text: 'A', list: 'bullet' }])
  doc.appendParagraphs([{ text: 'B', list: 'bullet' }])
  const ids = numIds(doc)
  t.is(ids[0], ids[1])
})

test('a document with no numbering part gains one', (t) => {
  const doc = docx.create()
  const files = unzip(doc.save())
  delete files['word/numbering.xml']
  const rels = decode(files['word/_rels/document.xml.rels']).replace(
    /<Relationship [^>]*numbering[^>]*\/>/,
    ''
  )
  const bare = docx.open(zipFiles({ ...files, 'word/_rels/document.xml.rels': utf8.encode(rels) }))
  bare.appendParagraphs([{ text: 'Bullet', list: 'bullet' }])
  const saved = unzip(bare.save())
  t.ok(saved['word/numbering.xml'])
  t.ok(decode(saved['word/_rels/document.xml.rels']).includes('numbering'))
  t.ok(decode(saved['[Content_Types].xml']).includes('numbering+xml'))
  t.is(para(reopen(bare).read().blocks[0]).list, true)
})

test('rejects an unknown list kind or level', (t) => {
  const doc = docx.create()
  t.exception(() => doc.appendParagraphs([JSON.parse('{"text":"x","list":"roman"}')]), /list must/)
  t.exception(() => doc.appendParagraphs([{ text: 'x', list: 'bullet', level: 9 }]), /level/)
})
