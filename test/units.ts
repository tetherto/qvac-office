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
  t.alike(units.lengths({ x: '1in', name: 'a' }, units.BOX_KEYS), { x: 914400, name: 'a' })
})

test('a malformed length names the field and what it expected', (t) => {
  try {
    units.length('two inches', 'width')
    t.fail('accepted a malformed length')
  } catch (err) {
    t.ok(err instanceof DocsError)
    const error = err as DocsError
    t.is(error.code, 'INVALID_BOX')
    t.is(error.details?.field, 'width')
    t.is(error.toJSON().details?.got, 'two inches')
  }
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
