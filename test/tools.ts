import test from 'brittle'
import b4a from 'b4a'
import { pptx, docx, xlsx, tools, DocsError } from '../lib/index.ts'
import { fixture } from './helpers/engine.ts'

const RED_PNG = fixture('red.png')

function base64(bytes: Uint8Array): string {
  return b4a.toString(bytes, 'base64')
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

test('every operation names a method of its format', (t) => {
  const docs = { pptx: pptx.create(), docx: docx.create(), xlsx: xlsx.create() }
  for (const format of ['pptx', 'docx', 'xlsx'] as const) {
    for (const op of tools.operations(format)) {
      const doc = docs[format] as unknown as Record<string, unknown>
      t.is(typeof doc[op.name], 'function', `${format}.${op.name}`)
      t.ok(op.description.length > 10, `${format}.${op.name} is described`)
    }
  }
})

test('a tool definition is one schema with every operation', (t) => {
  const definition = tools.tool('docx', { name: 'word' })
  t.is(definition.name, 'word')
  const items = (
    definition.input_schema as { properties: { ops: { items: { anyOf: unknown[] } } } }
  ).properties.ops.items.anyOf
  t.is(items.length, tools.operations('docx').length)
  t.is(tools.tool('xlsx', { only: ['read', 'setCells'] }).input_schema !== undefined, true)
  t.is(failure(() => tools.tool('xlsx', { only: ['nope'] })).code, 'INVALID_OPTIONS')
})

test('apply runs named operations and returns a new document', (t) => {
  const original = docx.create()
  const { doc, results } = tools.apply(original, [
    { op: 'appendParagraphs', paragraphs: [{ text: 'Plan', style: 'Heading 1' }, 'Body'] },
    { op: 'addTable', at: 2, rows: [['a', 'b']], header: true },
    { op: 'sections' }
  ])
  t.is(results.length, 3)
  t.ok(doc instanceof docx.Document)
  if (!(doc instanceof docx.Document)) {
    return
  }

  t.alike(
    (results[2] as { title: string }[]).map((s) => s.title),
    ['Plan']
  )
  t.ok(doc.read().blocks.some((block) => block.type === 'table'))
  t.absent(
    original.read().blocks.some((block) => block.type === 'table'),
    'the input is untouched'
  )
})

test('a failing op applies nothing and says which op failed', (t) => {
  const deck = pptx.create()
  deck.addSlide({ title: 'Hello' })
  const error = failure(() =>
    tools.apply(deck, [
      { op: 'setText', slide: 0, shape: 'title', text: 'Changed' },
      { op: 'deleteSlide', slide: 9 }
    ])
  )
  t.is(error.code, 'SLIDE_NOT_FOUND')
  t.is(error.details?.index, 1)
  t.is(error.details?.op, 'deleteSlide')
  t.ok(error.message.startsWith('ops[1] deleteSlide:'))
  const title = deck.read()[0]?.shapes.find((shape) => shape.placeholder === 'title')
  t.is(title?.paragraphs?.[0]?.text, 'Hello')
})

test('unknown operations and keys name what exists', (t) => {
  const book = xlsx.create()
  const unknown = failure(() => tools.apply(book, { op: 'explode' }))
  t.is(unknown.code, 'UNKNOWN_OPERATION')
  t.ok((unknown.details?.candidates as string[]).includes('setCells'))

  const extra = failure(() =>
    tools.apply(book, { op: 'setCells', sheet: 0, cells: {}, colour: 'red' })
  )
  t.is(extra.code, 'INVALID_OPTIONS')
  t.is(extra.details?.field, 'colour')

  const missing = failure(() => tools.apply(book, { op: 'setCells', sheet: 0 }))
  t.is(missing.details?.field, 'cells')
})

test('images arrive as base64 or a data URL', (t) => {
  const deck = pptx.create()
  deck.addSlide({})
  const { results } = tools.apply(deck, [
    { op: 'addPicture', slide: 0, image: base64(RED_PNG), width: '1in' },
    { op: 'addPicture', slide: 0, image: `data:image/png;base64,${base64(RED_PNG)}`, x: '2in' }
  ])
  t.is((results[0] as { width: number }).width, 914400)
  t.is(
    failure(() => tools.apply(deck, { op: 'addPicture', slide: 0, image: '!!' })).code,
    'INVALID_IMAGE'
  )
})

test('detect and open read the format from the bytes', (t) => {
  t.is(tools.detect(pptx.create().save()), 'pptx')
  t.is(tools.detect(docx.create().save()), 'docx')
  t.is(tools.detect(xlsx.create().save()), 'xlsx')
  t.ok(tools.open(xlsx.create().save()) instanceof xlsx.Workbook)
})

test('errors serialise for a tool result', (t) => {
  const error = failure(() => tools.apply(xlsx.create(), { op: 'read', sheet: 'Missing' }))
  const json = JSON.parse(JSON.stringify(error))
  t.is(json.code, 'SHEET_NOT_FOUND')
  t.is(json.details.op, 'read')
})
