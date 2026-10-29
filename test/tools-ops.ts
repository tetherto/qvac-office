import test from 'brittle'
import { docx, pptx, xlsx, tools } from '../lib/index.ts'
import { fixture } from './helpers/engine.ts'
import type { Operation } from '../lib/tools/index.ts'

function run(doc: tools.Doc, ops: Operation[]): unknown[] {
  return tools.apply(doc, ops).results
}

test('word operations run through apply', (t) => {
  const results = run(docx.create(), [
    { op: 'appendMarkdown', markdown: '# Plan\n\nRevenue grew **fast**.\n\n- one\n- two' },
    { op: 'appendParagraphs', paragraphs: [{ text: 'Step', list: 'number' }] },
    { op: 'replaceText', find: 'grew', replace: 'rose' },
    { op: 'formatText', index: 1, find: 'Revenue', format: { italic: true } },
    { op: 'link', index: 1, find: 'rose', url: 'https://example.com' },
    { op: 'addComment', index: 1, text: 'Check', find: 'Revenue' },
    { op: 'addFootnote', index: 1, text: 'Source' },
    { op: 'setPage', orientation: 'landscape', margins: { top: '1in' } },
    { op: 'addTableOfContents', at: 0, levels: 2 },
    { op: 'outline' },
    { op: 'find', heading: true },
    { op: 'read', maxChars: 200, detail: 'runs' },
    { op: 'markdown' },
    { op: 'page' }
  ])
  t.alike(results[2], { replaced: 1, blocks: [1] })
  t.is((results[9] as { headings: { title: string }[] }).headings.at(-1)?.title, 'Plan')
  t.ok(String((results[12] as { markdown: string }).markdown).includes('# Plan'))
  t.is((results[13] as { orientation: string }).orientation, 'landscape')
})

test('word block refs take ids from earlier results', (t) => {
  const doc = docx.create()
  const first = tools.apply(doc, { op: 'appendParagraphs', paragraphs: ['A', 'B'] })
  const [, id] = (first.results[0] as { ids: string[] }).ids
  const second = tools.apply(first.doc, [
    { op: 'insertParagraphs', at: { id }, paragraphs: 'Between' },
    { op: 'setText', index: { id }, text: 'B2' },
    { op: 'read' }
  ])
  const blocks = (second.results[2] as docx.ReadResult).blocks.map((block) =>
    block.type === 'table' ? '' : block.text
  )
  t.alike(blocks, ['A', 'Between', 'B2'])
})

test('powerpoint layout operations run through apply', (t) => {
  const deck = pptx.create()
  const results = run(deck, [
    { op: 'addSlide', title: 'Hi', body: 'Point' },
    { op: 'addTextBox', slide: 0, x: '1in', y: '1in', width: '2in', height: '1in', text: 'Box' },
    { op: 'setFill', slide: 0, shape: 4, color: 'FFF2CC' },
    { op: 'setLine', slide: 0, shape: 4, line: { color: '1F4E79', width: '2pt' } },
    { op: 'duplicateShape', slide: 0, shape: 4, to: 0, x: '4in' },
    { op: 'group', slide: 0, shapes: [4, 5], name: 'Pair' },
    { op: 'reorder', slide: 0, shape: 'title', order: 'front' },
    {
      op: 'addChart',
      slide: 0,
      into: 'body',
      replace: true,
      type: 'pie',
      categories: ['a', 'b'],
      series: [{ name: 's', values: [1, 2] }]
    },
    { op: 'setTheme', colors: { accent1: 'C00000' }, fonts: { heading: 'Georgia' } },
    { op: 'theme' },
    { op: 'addSlide', layout: 'Title and Content' },
    { op: 'addTable', slide: 1, into: 'body', rows: [['a', 'b']] },
    { op: 'setSize', size: '4:3', scale: true },
    { op: 'size' }
  ])
  t.is((results[4] as { x: number }).x, 4 * 914400)
  t.alike((results[5] as { children: number[] }).children.length, 2)
  t.is((results[9] as { colors: { accent1: string } }).colors.accent1, 'C00000')
  t.not((results[11] as { replaced?: number }).replaced, undefined)
  t.is((results[13] as { preset: string }).preset, '4:3')
})

test('excel calculation and templates run through apply', (t) => {
  const results = run(xlsx.create(), [
    { op: 'setCells', sheet: 0, cells: { A1: 2, A2: 3, A3: '=SUM(A1:A2)*2' } },
    { op: 'get', sheet: 0, ref: 'A3' },
    { op: 'calculate' },
    { op: 'toDocument' }
  ])
  t.is((results[1] as { value: number }).value, 10)
  t.is((results[3] as { changed: boolean }).changed, false)
})

function runsOf(doc: tools.Doc, index: number): unknown {
  if (!(doc instanceof docx.Document)) {
    throw new Error('expected a Word document')
  }

  const block = docx.open(doc.save()).read({ from: index, count: 1, detail: 'runs' }).blocks[0]

  return block.type === 'paragraph' ? block.runs : null
}

test('word offset editing runs through apply, one op per keystroke', (t) => {
  const { doc, results } = tools.apply(docx.open(fixture('report.docx')), [
    { op: 'insertText', index: 3, offset: 0, text: 'Q3: ', format: { italic: true } },
    { op: 'deleteText', index: 3, start: 4, end: 12 },
    { op: 'formatRange', index: 3, start: 0, end: 3, format: { bold: true } },
    { op: 'runFormat', index: 3, offset: 2 },
    { op: 'splitParagraph', index: 3, offset: 4 },
    { op: 'mergeParagraphs', index: 3 },
    { op: 'read', from: 3, count: 1 }
  ])
  t.alike(results[0], { length: 47, offset: 4 })
  t.alike(results[1], { length: 39, offset: 4 })
  t.alike(results[2], { changed: true })
  t.alike(results[3], { bold: true, italic: true, underline: false, size: 11 })
  const id = (results[6] as docx.ReadResult).blocks[0].id
  t.alike((results[4] as docx.SplitResult).first, { block: 3, id })
  t.alike(results[5], { block: 3, id, offset: 4, length: 39 })
  t.alike(runsOf(doc, 3), [
    { text: 'Q3:', bold: true, italic: true },
    { text: ' ', italic: true },
    { text: 'grew ' },
    { text: 'twelve percent', bold: true },
    { text: ' year over year.' }
  ])
  t.ok(tools.operations('docx', { only: ['runFormat'] })[0].query)
  t.absent(tools.operations('docx', { only: ['insertText'] })[0].query)
})
