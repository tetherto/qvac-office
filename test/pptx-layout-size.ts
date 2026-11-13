import test from 'brittle'
import * as pptx from '../lib/pptx/index.ts'
import { DocsError } from '../lib/errors.ts'
import { fixture, decode, unzip } from './helpers/engine.ts'

const reopen = (deck: pptx.Presentation): pptx.Presentation => pptx.open(deck.save())

function codeOf(fn: () => unknown): string | null {
  try {
    fn()
  } catch (err) {
    return err instanceof DocsError ? err.code : 'NOT_DOCS_ERROR'
  }

  return null
}

test('size reports the slide size and its preset', (t) => {
  t.alike(pptx.open(fixture('basic.pptx')).size(), {
    width: 9144000,
    height: 6858000,
    preset: '4:3'
  })
  t.alike(pptx.create().size(), { width: 12192000, height: 6858000, preset: '16:9' })
  t.alike(pptx.create({ size: { width: 914400 * 8, height: 914400 * 5 } }).size(), {
    width: 914400 * 8,
    height: 914400 * 5
  })
})

test('setSize changes the slide size and leaves shapes alone by default', (t) => {
  const deck = pptx.open(fixture('basic.pptx'))
  const before = deck.read()[2].shapes[1].box
  t.alike(deck.setSize('16:9'), {
    width: 12192000,
    height: 6858000,
    preset: '16:9',
    scaled: 0
  })
  const again = reopen(deck)
  t.alike(again.size().preset, '16:9')
  t.alike(again.read()[2].shapes[1].box, before)
  const files = unzip(deck.save())
  t.absent(decode(files['ppt/presentation.xml']).includes('screen4x3'))
  t.ok(files['ppt/slides/slide1.xml'])
})

test('setSize with scale stretches every shape, layout and master', (t) => {
  const deck = pptx.open(fixture('basic.pptx'))
  const result = deck.setSize({ width: '20in', height: '7.5in' }, { scale: true })
  t.is(result.width, 914400 * 20)
  t.ok(result.scaled > 0)
  const again = reopen(deck)
  t.alike(again.read()[2].shapes[1].box, {
    x: 914400 * 2,
    y: 1828800,
    width: 1828800 * 2,
    height: 914400
  })
  t.is(again.read()[1].shapes[0].box?.width, 8229600 * 2, 'a placeholder scales through its layout')
  const table = decode(unzip(deck.save())['ppt/slides/slide3.xml'])
  t.ok(table.includes('<a:gridCol w="3657600"/>'), 'table columns widen with the slide')
})

test('setSize refuses sizes out of range', (t) => {
  const deck = pptx.open(fixture('basic.pptx'))
  t.is(
    codeOf(() => deck.setSize({ width: 100, height: 100 })),
    'INVALID_OPTIONS'
  )
  t.is(
    codeOf(() => deck.setSize(JSON.parse('"5:4"'))),
    'INVALID_OPTIONS'
  )
  t.is(
    codeOf(() => deck.setSize({ width: 'wide', height: '7in' })),
    'INVALID_BOX'
  )
  t.is(deck.changed, false)
})
