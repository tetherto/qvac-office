import test from 'brittle'
import { pptx, docx, xlsx, units, DocsError } from '../lib/index.ts'
import { fixture } from './helpers/engine.ts'

test('lengths take EMU or a unit string', (t) => {
  t.is(units.length('2in'), 2 * 914400)
  t.is(units.length('5cm'), 5 * 360000)
  t.is(units.length('40 mm'), 40 * 36000)
  t.is(units.length('72pt'), 914400)
  t.is(units.length('96px'), 914400)
  t.is(units.length(12345), 12345)
  t.alike(units.box({ x: '1in', width: 5 }), {
    x: 914400,
    y: undefined,
    width: 5,
    height: undefined
  })
  t.alike(units.size({ height: '1cm' }), { width: undefined, height: 360000 })
})

function failure(fn: () => unknown): DocsError | null {
  try {
    fn()
  } catch (err) {
    return err instanceof DocsError ? err : null
  }

  return null
}

test('a malformed length names the field and what it expected', (t) => {
  const words = failure(() => units.length('two inches', 'width'))
  t.is(words?.code, 'INVALID_BOX')
  t.is(words?.details?.field, 'width')
  t.is(words?.toJSON().details?.got, 'two inches')
  t.is(failure(() => units.length(true, 'x'))?.details?.field, 'x')
})

test('a length outside the coordinate range is refused', (t) => {
  t.is(units.length(-27273042329600, 'x'), -27273042329600)
  t.is(units.length(27273042316900, 'width'), 27273042316900)
  t.is(units.length('-2in', 'x'), -2 * 914400)
  for (const value of ['1000000000000000in', 1e21, -27273042329601, Infinity, NaN]) {
    const error = failure(() => units.length(value, 'width'))
    t.is(error?.code, 'INVALID_BOX', String(value))
    t.is(error?.details?.field, 'width')
  }

  const deck = pptx.create()
  deck.addSlide({})
  const { id } = deck.addTextBox(0, { text: 'hi' })
  t.is(
    failure(() => deck.setBox({ slide: 0, shape: id }, { width: '1000000000000000in' }))?.code,
    'INVALID_BOX'
  )
  t.is(failure(() => deck.setBox({ slide: 0, shape: id }, { width: 1e21 }))?.code, 'INVALID_BOX')
})

test('every format places shapes with unit strings', (t) => {
  const deck = pptx.create()
  deck.addSlide({})
  const box = deck.addTextBox(0, { x: '1in', y: '1in', width: '3in', height: '1in', text: 'hi' })
  t.alike([box.x, box.width], [914400, 3 * 914400])

  const doc = docx.create()
  const image = doc.addImage(0, fixture('red.png'), { width: '1in' })
  t.is(image.width, 914400)

  const book = xlsx.create()
  const placed = book.addImage(0, fixture('red.png'), { width: '2cm', height: '2cm' })
  t.is(placed.width, 720000)
})
