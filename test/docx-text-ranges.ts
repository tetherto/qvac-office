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

function runsAt(doc: Document, index: number): ParagraphBlock['runs'] {
  return para(doc.read({ from: index, count: 1, detail: 'runs' }).blocks[0]).runs
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

const run = (value: string, rPr = ''): string =>
  `<w:r>${rPr}<w:t xml:space="preserve">${value}</w:t></w:r>`

test('replaceText matches across runs and keeps the formatting where the match starts', (t) => {
  const doc = open()
  t.alike(doc.replaceText({ find: 'grew twelve', replace: 'rose eleven' }), {
    replaced: 1,
    blocks: [3]
  })
  t.alike(runsAt(reopen(doc), 3), [
    { text: 'Revenue rose eleven' },
    { text: ' percent', bold: true },
    { text: ' year over year.' }
  ])
  const bold = open()
  bold.replaceText({ find: 'percent year', replace: 'points a' })
  t.alike(runsAt(reopen(bold), 3), [
    { text: 'Revenue grew ' },
    { text: 'twelve points a', bold: true },
    { text: ' over year.' }
  ])
})

test('replaceText replaces all matches, honours matchCase and limits to one block', (t) => {
  const doc = open()
  t.alike(doc.replaceText({ find: 'revenue', replace: 'Income', all: true }), {
    replaced: 2,
    blocks: [2, 3]
  })
  t.alike(doc.replaceText({ find: 'income', replace: 'x', matchCase: true }), {
    replaced: 0,
    blocks: []
  })
  t.alike(doc.replaceText({ find: 'year', replace: 'yr', all: true, in: 3 }), {
    replaced: 2,
    blocks: [3]
  })
  t.alike(doc.replaceText({ find: '25%', replace: '26%', in: 10 }), { replaced: 1, blocks: [10] })
  const read = reopen(doc).read().blocks
  t.is(para(read[3]).text, 'Income grew twelve percent yr over yr.')
  t.is(read[10].type === 'table' && read[10].rows[1][1], '26%')
})

test('replaceText takes a block id, tabs and line breaks, and an empty replacement', (t) => {
  const doc = docx.create()
  const { ids } = doc.appendParagraphs(['alpha beta gamma', 'alpha'])
  doc.replaceText({ find: 'beta', replace: 'b\tc\nd', in: { id: ids[0] } })
  doc.replaceText({ find: ' gamma', replace: '', in: { id: ids[0] } })
  const read = reopen(doc).read().blocks
  t.is(para(read[0]).text, 'alpha b\tc\nd')
  t.is(para(read[1]).text, 'alpha')
  t.is(read[0].id, ids[0])
})

test('a replace that finds nothing leaves the document unchanged', (t) => {
  const doc = open()
  t.alike(doc.replaceText({ find: 'absent', replace: 'x', all: true }), { replaced: 0, blocks: [] })
  t.is(doc.changed, false)
})

test('replaceText refuses tracked changes, fields and link boundaries before editing', (t) => {
  const tracked = withBody(
    `<w:p>${run('safe ')}<w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">${run('added')}</w:ins></w:p>`
  )
  const insert = failure(() => tracked.replaceText({ find: 'added', replace: 'x' }))
  t.is(insert.code, 'TRACKED_CHANGES')
  t.alike(insert.details, { block: 0, text: 'added' })
  const deletion = withBody(
    `<w:p>${run('one ')}<w:del w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del>${run('two')}</w:p>`
  )
  t.is(
    failure(() => deletion.replaceText({ find: 'one two', replace: 'x' })).code,
    'TRACKED_CHANGES'
  )
  const simple = withBody(
    `<w:p>${run('Page ')}<w:fldSimple w:instr="PAGE">${run('1')}</w:fldSimple></w:p>`
  )
  t.is(failure(() => simple.replaceText({ find: 'Page 1', replace: 'x' })).code, 'FIELD_TEXT')
  const complex = withBody(
    `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>DATE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${run('today')}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`
  )
  t.is(failure(() => complex.replaceText({ find: 'today', replace: 'x' })).code, 'FIELD_TEXT')
  const linked = withBody(
    `<w:p>${run('see ')}<w:hyperlink w:anchor="top">${run('the top')}</w:hyperlink></w:p>`
  )
  t.is(failure(() => linked.replaceText({ find: 'see the', replace: 'x' })).code, 'RANGE_BOUNDARY')
  t.alike(linked.replaceText({ find: 'top', replace: 'start' }), { replaced: 1, blocks: [0] })
  t.alike(runsAt(linked, 0), [{ text: 'see ' }, { text: 'the start', link: '#top' }])
  const mixed = withBody(
    `<w:p>${run('fine')}</w:p><w:p>${run('Page ')}<w:fldSimple w:instr="PAGE">${run('fine')}</w:fldSimple></w:p>`
  )
  t.is(
    failure(() => mixed.replaceText({ find: 'fine', replace: 'x', all: true })).code,
    'FIELD_TEXT'
  )
  t.is(para(mixed.read().blocks[0]).text, 'fine', 'nothing was replaced')
  t.is(mixed.changed, false)
})

test('replaceText validates its options', (t) => {
  const doc = open()
  t.exception(() => doc.replaceText({ find: '', replace: 'x' }), /options.find must be/)
  t.exception(() => Reflect.apply(doc.replaceText, doc, [{ find: 'a' }]), /options.replace must be/)
  t.is(failure(() => doc.replaceText({ find: 'a', replace: 'b', in: 99 })).code, 'BLOCK_NOT_FOUND')
})

test('formatText formats only the matched text by splitting runs', (t) => {
  const doc = open()
  t.alike(doc.formatText(3, { find: 'grew' }, { italic: true }), { formatted: 1 })
  t.alike(doc.formatText(3, { find: 'percent year', all: true }, { underline: true }), {
    formatted: 1
  })
  t.alike(runsAt(reopen(doc), 3), [
    { text: 'Revenue ' },
    { text: 'grew', italic: true },
    { text: ' ' },
    { text: 'twelve ', bold: true },
    { text: 'percent', bold: true, underline: true },
    { text: ' year', underline: true },
    { text: ' over year.' }
  ])
  t.is(para(reopen(doc).read().blocks[3]).text, 'Revenue grew twelve percent year over year.')
  t.is(failure(() => doc.formatText(3, { find: 'nothing' }, { bold: true })).code, 'TEXT_NOT_FOUND')
  t.exception(
    () => Reflect.apply(doc.formatText, doc, [3, { find: 'grew' }, { align: 'center' }]),
    /unknown format "align"/
  )
})

test('link wraps the match in an external hyperlink', (t) => {
  const doc = open()
  t.alike(doc.link(2, { find: 'revenue' }, 'mailto:finance@example.com'), { linked: 1 })
  const saved = doc.save()
  const files = unzip(saved)
  const rels = decode(files['word/_rels/document.xml.rels'])
  t.ok(rels.includes('Target="mailto:finance@example.com" TargetMode="External"'))
  t.ok(decode(files['word/document.xml']).includes('<w:hyperlink'))
  const again = docx.open(saved)
  t.is(runsAt(again, 2)?.[0].link, 'mailto:finance@example.com')
  t.is(
    failure(() => again.link(2, { find: 'Revenue' }, 'https://example.com')).code,
    'ALREADY_LINKED'
  )
  t.is(
    failure(() => again.link(2, { find: 'absent' }, 'https://example.com')).code,
    'TEXT_NOT_FOUND'
  )
})

test('link accepts only http, https and mailto URLs', (t) => {
  const doc = open()
  for (const url of ['javascript:alert(1)', 'file:///etc/passwd', 'https://a b', '', 'mailto:']) {
    const error = failure(() => doc.link(3, { find: 'grew' }, url))
    t.is(error.code, 'INVALID_LINK')
    t.is(error.details?.got, url)
  }

  t.is(doc.changed, false)
})

test('deleted tracked text takes no offsets, so matches after it land in place', (t) => {
  const doc = withBody(
    `<w:p>${run('keep')}<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:tab/><w:delText>gone</w:delText></w:r></w:del>${run(' more')}</w:p>`
  )
  t.alike(doc.replaceText({ find: 'more', replace: 'less' }), { replaced: 1, blocks: [0] })
  const block = para(reopen(doc).read({ detail: 'runs' }).blocks[0])
  t.is(block.text, 'keep less')
  t.is(block.runs?.map((r) => r.text).join(''), 'keep less')
})
