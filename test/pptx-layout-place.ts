import test from 'brittle'
import * as pptx from '../lib/pptx/index.ts'
import * as units from '../lib/units.ts'
import { DocsError } from '../lib/errors.ts'
import { fixture } from './helpers/engine.ts'

const reopen = (deck: pptx.Presentation): pptx.Presentation => pptx.open(deck.save())
const chart = {
  type: 'column' as const,
  categories: ['A', 'B'],
  series: [{ name: 'S', values: [1, 2] }]
}

function codeOf(fn: () => unknown): string | null {
  try {
    fn()
  } catch (err) {
    return err instanceof DocsError ? err.code : 'NOT_DOCS_ERROR'
  }

  return null
}

test('a chart with no box reports the body text it covers', (t) => {
  const deck = pptx.create()
  deck.addSlide({ title: 'Sales', body: ['Up and to the right'] })
  const placed = deck.addChart(0, chart)
  t.alike(placed.overlaps, [3], 'the body sits under the default box, the title above it')
})

test('overlaps skip empty placeholders, hidden shapes and full-slide pictures', (t) => {
  const deck = pptx.create()
  deck.addSlide({ title: 'Only a title' })
  const size = deck.size()
  const background = deck.addPicture(0, fixture('red.png'), {
    x: 0,
    y: 0,
    width: size.width,
    height: size.height
  })
  t.alike(background.overlaps, [2], 'the title, not the empty body')
  const box = deck.addTextBox(0, { text: 'Note', x: '1in', y: '3in', width: '2in', height: '1in' })
  t.alike(box.overlaps, [], 'the background picture does not count')
  const near = deck.addTextBox(0, {
    text: 'Near',
    x: '2in',
    y: '3.5in',
    width: '2in',
    height: '1in'
  })
  t.alike(near.overlaps, [box.id])
  const moved = deck.setBox({ slide: 0, shape: near.id }, { x: '6in' })
  t.alike(moved.overlaps, [])
})

test('into fills an empty placeholder with a chart, a table or a picture', (t) => {
  const deck = pptx.create()
  deck.addSlide({ title: 'Chart' })
  deck.addSlide({ title: 'Table' })
  deck.addSlide({ layout: 'Two Content', title: 'Picture' })
  const body = deck.read()[0].shapes[1]
  const placed = deck.addChart(0, { ...chart, into: 'body' })
  t.is(placed.replaced, body.id)
  const { x, y, width, height } = placed
  t.alike({ x, y, width, height }, body.box, 'the chart takes the layout box of the placeholder')
  t.alike(placed.overlaps, [])

  const table = deck.addTable(1, { rows: [['a', 'b']], into: 'body' })
  t.is(table.replaced, 3)

  const right = deck.read()[2].shapes[2]
  const picture = deck.addPicture(2, fixture('red.png'), { into: right.id ?? 0 })
  t.is(picture.replaced, right.id)
  const box = right.box
  t.ok(box && picture.x >= box.x && picture.x + picture.width <= box.x + box.width)
  t.ok(box && (picture.width === box.width || picture.height === box.height), 'it fills the box')

  const again = reopen(deck).read()
  t.alike(
    again.map((s) => s.shapes.map((shape) => shape.kind)),
    [
      ['text', 'chart'],
      ['text', 'table'],
      ['text', 'text', 'picture']
    ]
  )
})

test('into refuses a placeholder with text unless replace is set', (t) => {
  const deck = pptx.open(fixture('basic.pptx'))
  t.is(
    codeOf(() => deck.addChart(1, { ...chart, into: 'body' })),
    'PLACEHOLDER_HAS_TEXT'
  )
  t.is(
    codeOf(() => deck.addTable(2, { rows: [['x']], into: 3 })),
    'NOT_PLACEHOLDER'
  )
  t.is(
    codeOf(() => deck.addTable(1, { rows: [['x']], into: 'body', width: units.inches(1) })),
    'INVALID_OPTIONS'
  )
  t.is(deck.read()[1].shapes.length, 2, 'nothing was added')
  const placed = deck.addTable(1, { rows: [['x']], into: 'body', replace: true })
  t.is(placed.replaced, 3)
  t.alike(
    reopen(deck)
      .read()[1]
      .shapes.map((s) => s.kind),
    ['text', 'table']
  )
})
