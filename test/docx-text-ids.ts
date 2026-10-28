import test from 'brittle'
import * as docx from '../lib/docx/index.ts'
import { DocsError } from '../lib/errors.ts'
import { fixture, decode, unzip } from './helpers/engine.ts'
import type { Block, Document } from '../lib/docx/index.ts'

const reopen = (doc: Document): Document => docx.open(doc.save())
const documentXml = (doc: Document): string => decode(unzip(doc.save())['word/document.xml'])
const idsOf = (doc: Document): (string | undefined)[] => doc.read().blocks.map((b) => b.id)
const textOf = (block: Block): string => (block.type === 'table' ? '[table]' : block.text)

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

test('inserted paragraphs get fresh paraIds that read reports and saves keep', (t) => {
  const doc = docx.create()
  const result = doc.appendParagraphs(['One', 'Two', { text: 'Three', style: 'Heading 1' }])
  t.is(result.ids.length, 3)
  t.ok(result.ids.every((id) => /^[0-9A-F]{8}$/.test(id) && parseInt(id, 16) < 0x80000000))
  t.is(new Set(result.ids).size, 3)
  t.alike(idsOf(doc), result.ids)
  t.alike(idsOf(reopen(doc)), result.ids)
  const xml = documentXml(doc)
  t.ok(xml.includes('xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'))
  t.ok(xml.includes('mc:Ignorable="w14"'))
  t.ok(xml.includes(`w14:paraId="${result.ids[0]}"`))
  t.ok(xml.includes('w14:textId='))
})

test('ids keep pointing at the same paragraphs across inserts, deletes and reopen', (t) => {
  const doc = docx.create()
  const [first, second, third] = doc.appendParagraphs(['Alpha', 'Beta', 'Gamma']).ids
  const again = reopen(doc)
  const inserted = again.insertParagraphs({ id: second }, ['Before beta'])
  t.is(inserted.index, 1)
  t.alike(again.setText({ id: third }, 'Gamma two'), { changed: true })
  again.deleteBlocks({ id: first })
  const final = reopen(again)
  t.alike(final.read().blocks.map(textOf), ['Before beta', 'Beta', 'Gamma two'])
  t.alike(idsOf(final), [inserted.ids[0], second, third])
  final.formatParagraph({ id: second }, { bold: true })
  t.is(final.find({ text: 'beta', matchCase: true })[0].id, inserted.ids[0])
})

test('tables, rows, images and charts are addressed by their first paragraph id', (t) => {
  const doc = docx.create()
  doc.appendParagraphs(['Intro'])
  const table = doc.addTable(1, {
    rows: [
      ['A', 'B'],
      ['1', '2']
    ]
  })
  t.is(doc.read().blocks[1].id, table.id)
  doc.setCell({ table: { id: table.id }, row: 1, column: 0 }, 'one')
  doc.addRow({ id: table.id }, ['x', 'y'])
  doc.deleteRow({ table: { id: table.id }, row: 1 })
  const image = doc.addImage({ id: table.id }, fixture('red.png'))
  t.is(image.index, 1)
  t.is(doc.read().blocks[1].id, image.id)
  const xml = documentXml(doc)
  const all = [...xml.matchAll(/w14:paraId="([0-9A-F]{8})"/g)].map((m) => m[1])
  t.is(new Set(all).size, all.length, 'every paragraph id is unique, cloned rows included')
  const reopened = reopen(doc)
  const block = reopened.read().blocks[2]
  t.is(block.type, 'table')
  t.is(block.id, table.id)
  if (block.type === 'table') {
    t.alike(block.rows, [
      ['A', 'B'],
      ['x', 'y']
    ])
  }
})

test('an existing mc:Ignorable list is kept and not repeated', (t) => {
  const doc = docx.open(fixture('report.docx'))
  t.absent(doc.read().blocks[0].id)
  doc.insertParagraphs(0, ['First'])
  doc.insertParagraphs(0, ['Zeroth'])
  const xml = documentXml(doc)
  t.ok(xml.includes('mc:Ignorable="w14 wp14"'))
  t.is(xml.match(/xmlns:w14=/g)?.length, 1)
})

test('unknown or malformed ids fail with BLOCK_NOT_FOUND and details', (t) => {
  const doc = docx.create()
  const table = doc.addTable(0, { rows: [['A'], ['B']] })
  const unknown = failure(() => doc.setText({ id: '12345678' }, 'x'))
  t.is(unknown.code, 'BLOCK_NOT_FOUND')
  t.alike(unknown.details, { field: 'id', got: '12345678' })
  const malformed = failure(() => doc.deleteBlocks({ id: 'nope' }))
  t.is(malformed.code, 'BLOCK_NOT_FOUND')
  t.is(malformed.details?.got, 'nope')
  const inner = docx.open(doc.save())
  const xml = documentXml(inner)
  const second = [...xml.matchAll(/w14:paraId="([0-9A-F]{8})"/g)].map((m) => m[1])[1]
  const inside = failure(() => inner.setText({ id: second }, 'x'))
  t.is(inside.code, 'BLOCK_NOT_FOUND')
  t.is(inside.details?.inside, table.index)
})
