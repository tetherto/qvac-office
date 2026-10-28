import test from 'brittle'
import * as docx from '../lib/docx/index.ts'
import { fixture, decode, unzip } from './helpers/engine.ts'
import type { Block, Document, ParagraphBlock } from '../lib/docx/index.ts'

const open = (): Document => docx.open(fixture('report.docx'))
const reopen = (doc: Document): Document => docx.open(doc.save())

function para(block: Block | undefined): ParagraphBlock {
  if (!block || block.type !== 'paragraph') {
    throw new Error('expected a paragraph block')
  }

  return block
}

test('addComment anchors to a whole paragraph and creates comments.xml', (t) => {
  const doc = open()
  t.alike(doc.addComment(1, { text: 'Who exactly?', author: 'Ana', initials: 'AL' }), { id: '0' })
  const read = reopen(doc).read()
  t.alike(read.comments, [{ id: '0', author: 'Ana', text: 'Who exactly?' }])
  t.alike(para(read.blocks[1]).comments, ['0'])
  t.is(para(read.blocks[1]).text, 'Prepared by the finance team.')
  const files = unzip(doc.save())
  t.ok(decode(files['[Content_Types].xml']).includes('comments+xml'))
  t.ok(decode(files['word/comments.xml']).includes('w:initials="AL"'))
})

test('addComment with find splits runs around the found text', (t) => {
  const doc = open()
  doc.addComment(3, { text: 'Source?', find: 'grew twelve' })
  const body = decode(unzip(doc.save())['word/document.xml'])
  const start = body.indexOf('<w:commentRangeStart w:id="0"/>')
  const end = body.indexOf('<w:commentRangeEnd w:id="0"/>')
  const inside = body.slice(start, end)
  t.ok(start > 0 && end > start)
  t.ok(inside.includes('>grew <'))
  t.ok(inside.includes('>twelve<'))
  t.absent(inside.includes('percent'))
  t.is(para(reopen(doc).read().blocks[3]).text, 'Revenue grew twelve percent year over year.')
})

test('addComment reports text that is not there', (t) => {
  const doc = open()
  t.exception(() => doc.addComment(3, { text: 'x', find: 'shrank' }), /"shrank" is not in/)
  t.exception(() => doc.addComment(3, { text: '' }), /text must be non-empty/)
  t.exception(() => doc.addComment(10, { text: 'x' }), /not a paragraph/)
})

test('resolveComment marks a comment done in commentsExtended.xml', (t) => {
  const doc = open()
  const { id } = doc.addComment(1, { text: 'Done?' })
  t.alike(doc.resolveComment(id), { id, done: true, changed: true })
  t.alike(doc.resolveComment(id), { id, done: true, changed: false })
  const files = unzip(doc.save())
  const extended = decode(files['word/commentsExtended.xml'])
  t.ok(/w15:done="1"/.test(extended))
  const paraId = /w15:paraId="([0-9A-F]{8})"/.exec(extended)?.[1]
  t.ok(paraId && decode(files['word/comments.xml']).includes(`w14:paraId="${paraId}"`))
  t.ok(decode(files['word/_rels/document.xml.rels']).includes('commentsExtended'))
  const reopened = reopen(doc)
  t.alike(reopened.resolveComment(id, { done: false }), { id, done: false, changed: true })
})

test('deleteComment removes the comment, its anchors and its extension entry', (t) => {
  const doc = open()
  const first = doc.addComment(1, { text: 'One' })
  const second = doc.addComment(3, { text: 'Two', find: 'twelve' })
  doc.resolveComment(second.id)
  t.alike(doc.deleteComment(second.id), { id: second.id, anchors: 3 })
  const read = reopen(doc).read()
  t.alike(
    read.comments.map((c) => c.id),
    [first.id]
  )
  t.absent(para(read.blocks[3]).comments)
  t.absent(decode(unzip(doc.save())['word/commentsExtended.xml']).includes('commentEx '))
  t.exception(() => doc.deleteComment(7), /no comment with id 7; the comments are 0/)
})

test('addFootnote creates footnotes.xml with separators and places the reference', (t) => {
  const doc = open()
  t.alike(doc.addFootnote(3, { text: 'Audited figures.', find: 'twelve percent' }), { id: '1' })
  t.alike(doc.addFootnote(1, { text: 'Second note.' }), { id: '2' })
  const read = reopen(doc).read()
  t.alike(read.footnotes, [
    { id: '1', text: 'Audited figures.' },
    { id: '2', text: 'Second note.' }
  ])
  t.alike(para(read.blocks[3]).footnotes, ['1'])
  t.alike(para(read.blocks[1]).footnotes, ['2'])
  const files = unzip(doc.save())
  const notes = decode(files['word/footnotes.xml'])
  t.ok(notes.includes('w:type="separator" w:id="-1"'))
  t.ok(notes.includes('w:type="continuationSeparator" w:id="0"'))
  const body = decode(files['word/document.xml'])
  t.ok(body.indexOf('percent</w:t>') < body.indexOf('w:footnoteReference w:id="1"'))
  t.ok(body.indexOf('w:footnoteReference w:id="1"') < body.indexOf('year over year'))
})

test('comments and footnotes leave unrelated parts byte for byte', (t) => {
  const original = unzip(fixture('report.docx'))
  const doc = open()
  doc.addComment(1, { text: 'Note' })
  doc.addFootnote(1, { text: 'Foot' })
  const saved = unzip(doc.save())
  for (const name of ['word/styles.xml', 'word/numbering.xml', 'word/header1.xml']) {
    t.alike(saved[name], original[name], name)
  }
})
