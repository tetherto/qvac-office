import test from 'brittle'
import * as pptx from '../lib/pptx/index.ts'
import { DocsError } from '../lib/errors.ts'
import { fixture } from './helpers/engine.ts'
import { rich } from './helpers/pptx-text.ts'

const open = (): pptx.Presentation => pptx.open(fixture('basic.pptx'))
const reopen = (deck: pptx.Presentation): pptx.Presentation => pptx.open(deck.save())
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

function bad<T>(value: unknown): T {
  return value as T
}

test('read, addSlide and outline report stable slide ids', (t) => {
  const deck = open()
  t.alike(
    deck.read().map((s) => s.id),
    [256, 257, 258]
  )
  const added = deck.addSlide({ at: 0, title: 'First' })
  t.is(added.index, 0)
  t.is(added.id, 259)
  t.alike(
    reopen(deck)
      .outline()
      .map((s) => [s.index, s.id]),
    [
      [0, 259],
      [1, 256],
      [2, 257],
      [3, 258]
    ]
  )
})

test('every slide argument takes { id }', (t) => {
  const deck = open()
  deck.setText({ slide: { id: 258 }, shape: 'title' }, 'Data')
  t.alike(deck.moveSlide({ id: 258 }, { to: 0 }), { moved: true })
  t.is(deck.read({ slides: [{ id: 258 }] })[0].index, 0)
  deck.setNotes({ id: 256 }, 'Opening')
  deck.addTextBox({ id: 257 }, { text: 'Aside' })
  t.alike(deck.addSlide({ like: { id: 257 } }).layout, 'Title and Content')
  t.alike(deck.deleteSlide({ id: 257 }), { slides: 3 })

  const again = reopen(deck)
  t.alike(
    again.outline().map((s) => [s.id, s.title, s.notes]),
    [
      [258, 'Data', false],
      [256, 'Why the Sky Is Blue', true],
      [259, null, false]
    ]
  )
})

test('an unknown slide id is SLIDE_NOT_FOUND with the ids that exist', (t) => {
  const deck = open()
  const err = failure(() => deck.setText({ slide: { id: 999 }, shape: 'title' }, 'x'))
  t.is(err.code, 'SLIDE_NOT_FOUND')
  t.ok(err.message.includes('slide id 999 does not exist'))
  t.alike(err.details?.candidates, [256, 257, 258])
  t.is(failure(() => deck.moveSlide({ id: 1 }, { to: 0 })).code, 'SLIDE_NOT_FOUND')
  t.is(failure(() => deck.deleteSlide(bad({ id: '256' }))).code, 'SLIDE_NOT_FOUND')
  t.is(failure(() => deck.read({ slides: [{ id: 3 }] })).code, 'SLIDE_NOT_FOUND')
  t.alike(failure(() => deck.addPicture(7, fixture('red.png'))).details?.got, 7)
})

test('find matches text, placeholders, kinds and names', (t) => {
  const deck = open()
  t.alike(
    deck.find({ text: 'SCATTER' }).map((f) => [f.slide, f.shape]),
    [
      [0, 3],
      [1, 2],
      [1, 3]
    ]
  )
  t.alike(deck.find({ text: 'SCATTER', matchCase: true }), [])
  t.alike(
    deck.find({ text: '^Scat', regex: true }).map((f) => f.text),
    ['Scattering']
  )
  t.alike(
    deck.find({ placeholder: 'title' }).map((f) => f.slideId),
    [256, 257, 258]
  )
  t.alike(
    deck.find({ placeholder: 'subtitle' }).map((f) => f.name),
    ['Subtitle 2']
  )
  t.alike(deck.find({ kind: 'table' }), [
    {
      slide: 2,
      slideId: 258,
      shape: 4,
      name: 'Table 3',
      kind: 'table',
      text: 'Colour\nWavelength\nBlue\n450 nm'
    }
  ])
  t.alike(
    deck.find({ name: 'textbox', text: 'caption' }).map((f) => f.shape),
    [6]
  )
  t.alike(
    deck.find({ placeholder: 'body', kind: 'text' }).map((f) => f.shape),
    [3]
  )
})

test('find refuses unknown options, kinds and bad expressions', (t) => {
  const deck = open()
  t.is(failure(() => deck.find(bad({ style: 'x' }))).code, 'INVALID_OPTIONS')
  t.alike(failure(() => deck.find(bad({ kind: 'slide' }))).details?.field, 'kind')
  t.is(failure(() => deck.find({ text: '(', regex: true })).code, 'INVALID_OPTIONS')
  t.is(failure(() => deck.find(bad({ matchCase: 'yes' }))).code, 'INVALID_OPTIONS')
})

test('outline is small and counts what each slide holds', (t) => {
  const deck = pptx.create()
  for (let i = 0; i < 40; i++) {
    deck.addSlide({ title: `Slide ${i}`, body: 'x'.repeat(2000) })
  }

  const chart = deck.addSlide({ layout: 'Title Only', title: 'Chart' })
  deck.addChart(chart.index, {
    type: 'bar',
    categories: ['a', 'b'],
    series: [{ name: 's', values: [1, 2] }]
  })
  const outline = deck.outline()
  t.is(outline.length, 41)
  t.ok(JSON.stringify(outline).length < JSON.stringify(deck.read()).length / 10)
  t.alike(outline[40], {
    index: 40,
    id: chart.id,
    layout: 'Title Only',
    title: 'Chart',
    shapes: 2,
    notes: false,
    charts: 1,
    tables: 0,
    pictures: 0
  })
  t.alike(open().outline()[2], {
    index: 2,
    id: 258,
    layout: 'Title Only',
    title: 'Measurements',
    shapes: 5,
    notes: false,
    charts: 0,
    tables: 1,
    pictures: 1
  })
})

test('readPage continues from the same slide after a slide is added before it', (t) => {
  const deck = open()
  const sizes = deck.read().map((s) => JSON.stringify(s).length)
  const first = deck.readPage({ maxChars: sizes[0] })
  deck.addSlide({ at: 0, layout: 'Title Only', title: 'Inserted' })
  const rest = deck.readPage({ from: first.next })
  t.alike(
    rest.slides.map((s) => s.id),
    [257, 258]
  )
})

test('readPage stops before maxChars and hands back the next slide', (t) => {
  const deck = open()
  const sizes = deck.read().map((s) => JSON.stringify(s).length)
  const first = deck.readPage({ maxChars: sizes[0] + sizes[1] })
  t.alike(
    first.slides.map((s) => s.index),
    [0, 1]
  )
  t.alike(first.next, { id: 258 })
  const rest = deck.readPage({ from: first.next, maxChars: sizes[0] + sizes[1] })
  t.alike(
    rest.slides.map((s) => s.index),
    [2]
  )
  t.absent(rest.next)
  const tiny = deck.readPage({ maxChars: 1 })
  t.is(tiny.slides.length, 1)
  t.alike(tiny.next, { id: 257 })
  t.is(deck.readPage({ from: { id: 258 } }).slides[0].id, 258)
  t.is(deck.readPage().slides.length, 3)
  t.is(failure(() => deck.readPage({ maxChars: 0 })).code, 'INVALID_OPTIONS')
  t.is(failure(() => deck.read(bad({ maxChars: 10 }))).code, 'INVALID_OPTIONS')
  t.is(failure(() => deck.read(bad({ detail: 'words' }))).code, 'INVALID_OPTIONS')
})

test("read({ detail: 'runs' }) reports each run's own formatting, links and fields", (t) => {
  const deck = rich()
  deck.link({ slide: 1, shape: 3 }, { find: 'next' }, 'https://example.com/a')
  const shape = reopen(deck).read({ detail: 'runs' })[1].shapes[1]
  t.alike(shape.paragraphs?.[0], {
    text: 'Sunlight2\nnext',
    level: 0,
    runs: [
      { text: 'Sun', bold: true, size: 24, color: '1A73E8', font: 'Georgia' },
      { text: 'light', italic: true, underline: true },
      { text: '2', field: 'slidenum' },
      { text: '\n' },
      { text: 'next', link: 'https://example.com/a' }
    ]
  })
  t.alike(shape.paragraphs?.[1].runs, [{ text: 'Short wavelengths scatter more' }])
  t.absent(open().read()[1].shapes[1].paragraphs?.[0].runs)
  t.ok(deck.readPage({ detail: 'runs' }).slides[1].shapes[1].paragraphs?.[0].runs)
})
