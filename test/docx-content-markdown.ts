import test from 'brittle'
import * as docx from '../lib/docx/index.ts'
import { parse, inlines } from '../lib/docx/markdown.ts'
import { fixture, decode, unzip } from './helpers/engine.ts'
import type { Block, Document, ParagraphBlock, TableBlock } from '../lib/docx/index.ts'

const reopen = (doc: Document): Document => docx.open(doc.save())

const SAMPLE = `# Plan

Intro with **bold**, *italic*, \`code\` and a [link](https://example.com/a).

## Steps

- Gather
  - Sort *carefully*
- Ship
1. First
2. Second

| Name | Score |
| --- | --- |
| Ada | 3 |
| Bob \\| Co | 4 |

---

Last line.
`

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

test('parses inline bold, italic, code, links and escapes', (t) => {
  t.alike(inlines('a **b** *c* `d` [e](http://x) \\*f\\* ***g***'), [
    { text: 'a ' },
    { text: 'b', bold: true },
    { text: ' ' },
    { text: 'c', italic: true },
    { text: ' ' },
    { text: 'd', code: true },
    { text: ' ' },
    { text: 'e', link: 'http://x' },
    { text: ' *f* ' },
    { text: 'g', bold: true, italic: true }
  ])
  t.alike(inlines('snake_case_name'), [{ text: 'snake_case_name' }])
  t.alike(inlines('2 * 3 * 4'), [{ text: '2 * 3 * 4' }])
})

test('parses headings, nested lists, tables, rules and fences', (t) => {
  const blocks = parse(SAMPLE + '\n```\nraw *text*\n```\n')
  t.alike(
    blocks.map((b) => b.type),
    [
      'heading',
      'paragraph',
      'heading',
      'item',
      'item',
      'item',
      'item',
      'item',
      'table',
      'rule',
      'paragraph',
      'code'
    ]
  )
  const levels = blocks.flatMap((b) => (b.type === 'item' ? [[b.ordered, b.level]] : []))
  t.alike(levels, [
    [false, 0],
    [false, 1],
    [false, 0],
    [true, 0],
    [true, 0]
  ])
})

test('appendMarkdown writes headings, lists, tables and links', (t) => {
  const doc = docx.create()
  const result = doc.appendMarkdown(SAMPLE)
  t.alike({ ...result, ids: result.ids.length }, { index: 0, count: 11, ids: 11 })
  const blocks = reopen(doc).read().blocks
  t.is(para(blocks[0]).heading, 1)
  t.is(para(blocks[2]).style, 'Heading 2')
  t.alike(
    blocks.slice(3, 8).map((b) => [para(b).text, para(b).list, para(b).level]),
    [
      ['Gather', true, 0],
      ['Sort carefully', true, 1],
      ['Ship', true, 0],
      ['First', true, 0],
      ['Second', true, 0]
    ]
  )
  t.alike(table(blocks[8]).rows, [
    ['Name', 'Score'],
    ['Ada', '3'],
    ['Bob | Co', '4']
  ])
  const files = unzip(doc.save())
  const body = decode(files['word/document.xml'])
  t.ok(body.includes('<w:hyperlink'))
  t.ok(body.includes('<w:tblHeader/>'))
  t.ok(body.includes('w:ascii="Consolas"'))
  t.ok(body.includes('<w:pBdr>'))
  t.ok(
    decode(files['word/_rels/document.xml.rels']).includes(
      'Target="https://example.com/a" TargetMode="External"'
    )
  )
})

test('markdown() round-trips what appendMarkdown wrote', (t) => {
  const doc = docx.create()
  doc.appendMarkdown(SAMPLE)
  const first = reopen(doc).markdown()
  t.is(first.total, 11)
  t.ok(first.markdown.includes('[link](https://example.com/a)'))
  t.ok(first.markdown.includes('| Bob \\| Co | 4 |'))
  t.ok(first.markdown.includes('    - Sort *carefully*'))
  const again = docx.create()
  again.appendMarkdown(first.markdown)
  t.is(reopen(again).markdown().markdown, first.markdown)
})

test('insertMarkdown places blocks at an index and markdown() pages with from and count', (t) => {
  const doc = docx.open(fixture('report.docx'))
  const before = doc.blockCount
  const inserted = doc.insertMarkdown(2, 'Inserted **here**')
  t.alike({ ...inserted, ids: inserted.ids.length }, { index: 2, count: 1, ids: 1 })
  const reread = reopen(doc)
  t.is(reread.blockCount, before + 1)
  t.is(para(reread.read().blocks[2]).text, 'Inserted here')
  const page = reread.markdown({ from: 0, count: 4 })
  t.is(page.total, before + 1)
  t.is(
    page.markdown,
    '# Quarterly Report\n\nPrepared by the finance team.\n\nInserted **here**\n\n# Revenue\n'
  )
})

test('markdown escapes text that would read as markup', (t) => {
  const plain = ['# not a heading', '1. not a list', 'a*b*c', '- not an item', '---', '> quote']
  const doc = docx.create()
  doc.appendParagraphs([...plain, { runs: [{ text: 'Bold', bold: true }, ' start'] }])
  const out = doc.markdown().markdown
  t.is(
    out,
    '\\# not a heading\n\n1\\. not a list\n\na\\*b\\*c\n\n\\- not an item\n\n\\---\n\n\\> quote\n\n**Bold** start\n'
  )
  const again = docx.create()
  again.appendMarkdown(out)
  t.alike(
    again.read().blocks.map((b) => para(b).text),
    [...plain, 'Bold start']
  )
})

test('appendMarkdown refuses empty input', (t) => {
  t.exception(() => docx.create().appendMarkdown('  \n\n'), /no blocks/)
})
