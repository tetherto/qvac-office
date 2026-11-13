import test from 'brittle'
import * as pptx from '../lib/pptx/index.ts'
import { DocsError } from '../lib/errors.ts'
import { fixture, decode, unzip } from './helpers/engine.ts'
import { rich, withParagraph, SLIDE2 } from './helpers/pptx-text.ts'

const open = (): pptx.Presentation => pptx.open(fixture('basic.pptx'))
const reopen = (deck: pptx.Presentation): pptx.Presentation => pptx.open(deck.save())
const SPLIT =
  '<a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Reve</a:t></a:r><a:r><a:rPr lang="en-US" i="1"/><a:t>nue gr</a:t></a:r><a:r><a:rPr lang="en-US"/><a:t>ew fast</a:t></a:r></a:p>'

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

function runs(deck: pptx.Presentation, slide: number, shape: number): pptx.ReadRun[] {
  const found = deck.read({ detail: 'runs' })[slide].shapes.find((s) => s.id === shape)
  return found?.paragraphs?.[0].runs ?? []
}

test('replaceText matches across runs and keeps the formatting where the match starts', (t) => {
  const deck = withParagraph(SPLIT)
  t.alike(deck.replaceText({ find: 'revenue grew', replace: 'Sales rose' }), {
    replaced: 1,
    shapes: [{ slide: 1, shape: 3 }]
  })
  t.alike(runs(reopen(deck), 1, 3), [{ text: 'Sales rose', bold: true }, { text: ' fast' }])
})

test('replaceText replaces the first match, or every match with all', (t) => {
  const first = open()
  t.alike(first.replaceText({ find: 'scatter', replace: 'spread' }).shapes, [
    { slide: 0, shape: 3 }
  ])
  t.is(reopen(first).read()[1].shapes[1].paragraphs?.[1].text, 'Short wavelengths scatter more')

  const every = open()
  const result = every.replaceText({ find: 'scatter', replace: 'spread', all: true })
  t.is(result.replaced, 3)
  t.alike(result.shapes, [
    { slide: 0, shape: 3 },
    { slide: 1, shape: 2 },
    { slide: 1, shape: 3 }
  ])
  t.alike(
    reopen(every)
      .find({ text: 'spread' })
      .map((f) => f.text),
    [
      'A short tour of Rayleigh spreading',
      'spreading',
      'Sunlight holds every colour\nShort wavelengths spread more'
    ]
  )
  t.is(
    open().replaceText({ find: 'Scatter', replace: 'x', matchCase: true, all: true }).replaced,
    1
  )
  t.alike(open().replaceText({ find: 'nothing here', replace: 'x' }), { replaced: 0, shapes: [] })
})

test('replaceText scopes to a slide, a shape, table cells and notes', (t) => {
  const deck = open()
  t.is(
    deck.replaceText({ find: 'scatter', replace: 'x', all: true, slide: { id: 257 } }).replaced,
    2
  )
  t.is(deck.replaceText({ find: 'blue', replace: 'Azure', slide: 2, shape: 4 }).replaced, 1)
  t.is(deck.replaceText({ find: 'sunset', replace: 'dusk', all: true }).replaced, 0)
  t.alike(deck.replaceText({ find: 'sunset', replace: 'dusk', slide: 1, notes: true }), {
    replaced: 1,
    shapes: [{ slide: 1, shape: 'notes' }]
  })

  const again = reopen(deck)
  const slides = again.read()
  t.alike(slides[2].shapes[2].rows?.[1], ['Azure', '450 nm'])
  t.is(slides[0].shapes[1].paragraphs?.[0].text, 'A short tour of Rayleigh scattering')
  t.is(slides[1].notes, 'Mention the dusk question.')
  t.is(
    failure(() => deck.replaceText(bad({ find: 'x', replace: 'y', shape: 3 }))).code,
    'INVALID_OPTIONS'
  )
  t.is(
    failure(() => deck.replaceText({ find: 'x', replace: 'y', slide: 2, shape: 3 })).code,
    'NOT_TEXT'
  )
})

test('replaceText refuses matches that cross a field or a line break', (t) => {
  const deck = rich()
  t.alike(deck.replaceText({ find: 'light2', replace: 'x' }), {
    replaced: 0,
    shapes: [],
    refused: 1
  })
  t.alike(deck.replaceText({ find: '2', replace: 'x', slide: 1 }), {
    replaced: 0,
    shapes: [],
    refused: 1
  })
  t.is(
    failure(() => deck.replaceText({ find: 'light\nnext', replace: 'x' })).code,
    'INVALID_OPTIONS'
  )
  t.is(failure(() => deck.replaceText({ find: 'next', replace: 'a\nb' })).code, 'INVALID_OPTIONS')
  t.is(failure(() => deck.replaceText({ find: '', replace: 'b' })).code, 'INVALID_OPTIONS')
  t.is(deck.replaceText({ find: 'Sunlight', replace: 'Moonlight' }).replaced, 1)
  t.is(reopen(deck).read()[1].shapes[1].paragraphs?.[0].text, 'Moonlight2\nnext')
})

test('replaceText leaves other parts byte for byte', (t) => {
  const deck = open()
  deck.replaceText({ find: 'Scattering', replace: 'Spread', matchCase: true })
  const before = unzip(fixture('basic.pptx'))
  const after = unzip(deck.save())
  for (const name of Object.keys(before)) {
    if (name !== SLIDE2) {
      t.alike(after[name], before[name], name)
    }
  }
})

test('formatText styles only the matched text', (t) => {
  const deck = open()
  const result = deck.formatText(
    { slide: 1, shape: 3 },
    { find: 'every' },
    { italic: true, color: 'C0392B' }
  )
  t.alike(result, { formatted: 1, runs: 1 })
  t.alike(runs(reopen(deck), 1, 3), [
    { text: 'Sunlight holds ', bold: true, size: 24 },
    { text: 'every', bold: true, italic: true, size: 24, color: 'C0392B' },
    { text: ' colour', bold: true, size: 24 }
  ])

  const across = withParagraph(SPLIT)
  t.alike(across.formatText({ slide: 1, shape: 3 }, { find: 'venue g' }, { underline: true }), {
    formatted: 1,
    runs: 2
  })
  t.alike(runs(reopen(across), 1, 3), [
    { text: 'Re', bold: true },
    { text: 've', bold: true, underline: true },
    { text: 'nue g', italic: true, underline: true },
    { text: 'r', italic: true },
    { text: 'ew fast' }
  ])
})

test('formatText with all styles every match, and refuses what it cannot style', (t) => {
  const deck = open()
  t.alike(deck.formatText({ slide: 2, shape: 4 }, { find: 'l', all: true }, { bold: true }), {
    formatted: 3,
    runs: 3
  })
  const err = failure(() =>
    deck.formatText({ slide: 1, shape: 3 }, { find: 'purple' }, { bold: true })
  )
  t.is(err.code, 'TEXT_NOT_FOUND')
  t.is(err.details?.got, 'purple')
  t.is(
    failure(() => deck.formatText({ slide: 1, shape: 3 }, { find: 'every' }, { align: 'center' }))
      .code,
    'INVALID_STYLE'
  )
  t.is(
    failure(() => deck.formatText({ slide: 2, shape: 3 }, { find: 'x' }, { bold: true })).code,
    'NOT_TEXT'
  )
  t.alike(rich().formatText({ slide: 1, shape: 3 }, { find: 'light2' }, { bold: true }), {
    formatted: 0,
    runs: 0,
    refused: 1
  })
})

test('link adds an external hyperlink to the matched runs only', (t) => {
  const deck = open()
  t.alike(deck.link({ slide: 1, shape: 3 }, { find: 'colour' }, 'https://example.com/light'), {
    linked: 1,
    runs: 1
  })
  deck.link({ slide: 1, shape: 2 }, { find: 'Scatter' }, 'https://example.com/light')
  deck.link({ slide: 1, shape: 3 }, { find: 'Sunlight' }, 'mailto:sky@example.com')
  const files = unzip(deck.save())
  const rels = decode(files['ppt/slides/_rels/slide2.xml.rels'])
  t.is(rels.split('https://example.com/light').length, 2)
  t.ok(rels.includes('Target="mailto:sky@example.com" TargetMode="External"'))
  t.ok(decode(files[SLIDE2]).includes('<a:hlinkClick r:id='))

  const again = pptx.open(deck.save())
  t.alike(runs(again, 1, 3), [
    { text: 'Sunlight', bold: true, size: 24, link: 'mailto:sky@example.com' },
    { text: ' holds every ', bold: true, size: 24 },
    { text: 'colour', bold: true, size: 24, link: 'https://example.com/light' }
  ])
})

test('link takes only http, https and mailto URLs', (t) => {
  const deck = open()
  for (const url of [
    'javascript:alert(1)',
    'file:///etc/passwd',
    'example.com',
    'https://',
    'http://a b'
  ]) {
    const err = failure(() => deck.link({ slide: 1, shape: 3 }, { find: 'every' }, url))
    t.is(err.code, 'INVALID_LINK', url)
  }

  t.is(
    failure(() => deck.link({ slide: 1, shape: 3 }, { find: 'purple' }, 'https://x.io')).code,
    'TEXT_NOT_FOUND'
  )
  t.absent(deck.changed)
})

test('text edits survive deleting and moving the slides around them', (t) => {
  const deck = open()
  const [hit] = deck.find({ text: 'Measurements' })
  deck.deleteSlide(0)
  deck.formatText(
    { slide: { id: hit.slideId }, shape: hit.shape },
    { find: 'Measure' },
    { bold: true }
  )
  t.alike(runs(reopen(deck), 1, 2)[0], { text: 'Measure', bold: true })
})
