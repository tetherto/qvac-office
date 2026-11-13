import test from 'brittle'
import { docx, pptx, DocsError } from '../lib/index.ts'
import { render } from '../lib/render/index.ts'

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47]
const onBare = typeof (globalThis as { Bare?: unknown }).Bare !== 'undefined'

test('render draws pages to PNG on Bare', { skip: !onBare }, async (t) => {
  const deck = pptx.create()
  deck.addSlide({ title: 'One' })
  deck.addSlide({ title: 'Two' })
  const result = await render(deck, { pages: [1], scale: 0.5 })
  t.is(result.pageCount, 2)
  t.is(result.pages.length, 1)
  t.is(result.pages[0]?.page, 1)
  t.alike([...(result.pages[0]?.png.subarray(0, 4) ?? [])], PNG_SIGNATURE)
  await t.exception(render(deck, { pages: [5] }), /pages must be from 0 to 1/)
})

test('render says what it needs where it cannot run', { skip: onBare }, async (t) => {
  try {
    await render(docx.create())
    t.fail('rendered without Bare')
  } catch (err) {
    t.ok(err instanceof DocsError)
    t.is((err as DocsError).code, 'RENDER_UNAVAILABLE')
  }
})
