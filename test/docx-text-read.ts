import test from 'brittle'
import * as docx from '../lib/docx/index.ts'
import { fixture } from './helpers/engine.ts'
import type { Block, Document, ParagraphBlock } from '../lib/docx/index.ts'

const open = (): Document => docx.open(fixture('report.docx'))

function para(block: Block | undefined): ParagraphBlock {
  if (!block || block.type !== 'paragraph') {
    throw new Error('expected a paragraph block')
  }

  return block
}

test('find matches text, regex, style and heading, alone or together', (t) => {
  const doc = open()
  t.alike(
    doc.find({ text: 'REVENUE' }).map((f) => f.index),
    [2, 3]
  )
  t.alike(doc.find({ text: 'REVENUE', matchCase: true }), [])
  t.alike(doc.find({ regex: '^Revenue\\b', heading: false }), [
    {
      index: 3,
      type: 'paragraph',
      text: 'Revenue grew twelve percent year over year.',
      style: 'Normal'
    }
  ])
  t.alike(
    doc.find({ style: 'list bullet' }).map((f) => f.text),
    ['North led the growth', 'South held flat']
  )
  t.alike(
    doc.find({ heading: 1 }).map((f) => f.text),
    ['Revenue', 'Costs', 'Outlook']
  )
  t.alike(
    doc.find({ text: 'margin' }).map((f) => [f.index, f.type]),
    [[10, 'table']]
  )
})

test('find rejects empty, unknown and malformed queries', (t) => {
  const doc = open()
  t.exception(() => doc.find({}), /at least one of text, regex, style or heading/)
  t.exception(() => doc.find({ regex: '(' }), /regex must be a valid regular expression/)
  const loose: Record<string, unknown> = { placeholder: 'x' }
  t.exception(() => doc.find(loose), /query must be made of/)
  t.exception(() => doc.find({ heading: 12 }), /heading level from 0 to 9/)
})

test('outline lists headings with the block range under each', (t) => {
  const doc = open()
  const outline = doc.outline()
  t.is(outline.total, 14)
  t.ok(outline.chars > 100)
  t.alike(outline.headings, [
    { level: 1, title: 'Revenue', index: 2, end: 8 },
    { level: 2, title: 'Regions', index: 6, end: 8 },
    { level: 1, title: 'Costs', index: 8, end: 11 },
    { level: 1, title: 'Outlook', index: 11, end: 14 }
  ])
  t.alike(
    doc.outline({ depth: 1 }).headings.map((h) => h.title),
    ['Revenue', 'Costs', 'Outlook']
  )
  t.ok(JSON.stringify(outline).length < 600)
  t.exception(() => doc.outline({ depth: 0 }), /depth must be a heading level/)
})

test('outline reports ids for headings that have one', (t) => {
  const doc = docx.create()
  const { ids } = doc.appendParagraphs([{ text: 'Top', style: 'Heading 1' }, 'Body'])
  t.alike(doc.outline().headings, [{ level: 1, title: 'Top', index: 0, id: ids[0], end: 2 }])
})

test('read stops at maxChars, returns at least one block and a next cursor', (t) => {
  const doc = open()
  const first = doc.read({ maxChars: 1 })
  t.is(first.blocks.length, 1)
  t.is(first.next, 1)
  const pages: number[][] = []
  for (let from: number | undefined = 0; from !== undefined;) {
    const page = doc.read({ from, maxChars: 300 })
    t.ok(page.blocks.length === 1 || JSON.stringify(page.blocks).length <= 300 + page.blocks.length)
    pages.push(page.blocks.map((b) => b.index))
    from = page.next
  }

  t.ok(pages.length > 2)
  t.alike(pages.flat(), [...Array(14).keys()])
  t.is(doc.read({ count: 3 }).next, 3)
  t.absent(doc.read().next)
  t.exception(() => doc.read({ maxChars: 0 }), /maxChars must be a positive whole number/)
})

test('detail runs reports run formatting and links', (t) => {
  const doc = open()
  const read = doc.read({ from: 3, count: 1, detail: 'runs' })
  t.alike(para(read.blocks[0]).runs, [
    { text: 'Revenue grew ' },
    { text: 'twelve percent', bold: true },
    { text: ' year over year.' }
  ])
  t.absent(para(doc.read({ from: 3, count: 1 }).blocks[0]).runs)
  doc.link(3, { find: 'year over year' }, 'https://example.com/yoy')
  doc.formatText(
    3,
    { find: 'Revenue' },
    { italic: true, size: 14, color: 'C0392B', font: 'Georgia' }
  )
  const runs = para(
    docx.open(doc.save()).read({ from: 3, count: 1, detail: 'runs' }).blocks[0]
  ).runs
  t.alike(runs, [
    { text: 'Revenue', italic: true, size: 14, color: 'C0392B', font: 'Georgia' },
    { text: ' grew ' },
    { text: 'twelve percent', bold: true },
    { text: ' ' },
    { text: 'year over year', underline: true, color: '0563C1', link: 'https://example.com/yoy' },
    { text: '.' }
  ])
  t.exception(() => Reflect.apply(doc.read, doc, [{ detail: 'words' }]), /detail must be "runs"/)
})
