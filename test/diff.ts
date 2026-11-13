import test from 'brittle'
import { pptx, docx, xlsx, diff } from '../lib/index.ts'

test('a Word diff lists added, removed and changed blocks', (t) => {
  const before = docx.create()
  before.appendParagraphs(['One', 'Two', 'Three'])
  const after = docx.open(before.save())
  after.setText(1, 'Two, revised')
  after.deleteBlocks(2)
  after.appendParagraphs(['Four'])
  after.setHeader('Draft')

  const result = diff.diff(before.save(), after.save())
  t.is(result.format, 'docx')
  t.alike(result.counts, { added: 0, removed: 0, changed: 3 })
  t.ok(result.changes.some((c) => c.before === 'Two' && c.after === 'Two, revised'))
  t.ok(result.changes.some((c) => c.where === 'headers' && c.after === 'Draft'))
})

test('a PowerPoint diff reports slides and shapes', (t) => {
  const before = pptx.create()
  before.addSlide({ title: 'Hello' })
  const after = pptx.open(before.save())
  after.setText({ slide: 0, shape: 'title' }, 'Hello again')
  after.addSlide({ title: 'New' })
  after.setNotes(0, 'Say hi')

  const result = diff.diff(before, after)
  t.ok(
    result.changes.some(
      (c) => c.where.startsWith('slide 0 shape') && c.after?.includes('Hello again')
    )
  )
  t.ok(result.changes.some((c) => c.where === 'slide 0 notes'))
  t.ok(result.changes.some((c) => c.kind === 'added' && c.where === 'slide 1'))
})

test('an Excel diff lists cells and sheets, and truncates long lists', (t) => {
  const before = xlsx.create()
  before.setCells(0, { A1: 1, A2: 2 })
  const after = xlsx.open(before.save())
  after.setCells(0, { A1: 5, A3: '=A1+A2' })
  after.addSheet('Notes')

  const result = diff.diff(before.save(), after.save())
  t.ok(result.changes.some((c) => c.where === 'Sheet1!A1' && c.before === '1' && c.after === '5'))
  t.ok(result.changes.some((c) => c.where === 'Sheet1!A3' && c.kind === 'added'))
  t.ok(result.changes.some((c) => c.where === 'sheet Notes' && c.kind === 'added'))

  const many = xlsx.open(before.save())
  many.setCells(0, Object.fromEntries(Array.from({ length: 30 }, (_, i) => [`B${i + 1}`, i])))
  const capped = diff.diff(before.save(), many.save(), { limit: 10 })
  t.is(capped.changes.length, 10)
  t.is(capped.truncated, true)
  t.is(capped.counts.added, 30)
})

test('diff refuses two different formats', (t) => {
  t.exception(() => diff.diff(docx.create().save(), xlsx.create().save()), /one format/)
})

test('an Excel diff compares cells past the first ten thousand', (t) => {
  const before = xlsx.create()
  before.setCells(0, { A1: 1, AX300: 2, A290: 3 })
  const after = xlsx.open(before.save())
  after.setCells(0, { A290: 4 })

  const result = diff.diff(before.save(), after.save())
  t.alike(
    result.changes.map((c) => c.where),
    ['Sheet1!A290']
  )
  t.is(result.truncated, false)
})

test('an Excel diff says it is truncated when a sheet is too large to compare', (t) => {
  const before = xlsx.create()
  before.setCells(0, { A1: 1, Z50000: 2 })
  const after = xlsx.open(before.save())
  after.setCells(0, { Z50000: 3 })

  const result = diff.diff(before.save(), after.save())
  t.is(result.truncated, true)
})
