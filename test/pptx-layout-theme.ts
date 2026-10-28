import test from 'brittle'
import * as pptx from '../lib/pptx/index.ts'
import { DocsError } from '../lib/errors.ts'
import { fixture, decode, unzip } from './helpers/engine.ts'

const open = (): pptx.Presentation => pptx.open(fixture('basic.pptx'))
const reopen = (deck: pptx.Presentation): pptx.Presentation => pptx.open(deck.save())

function codeOf(fn: () => unknown): string | null {
  try {
    fn()
  } catch (err) {
    return err instanceof DocsError ? err.code : 'NOT_DOCS_ERROR'
  }

  return null
}

test('theme reads the colour scheme and fonts', (t) => {
  const theme = open().theme()
  t.is(theme.name, 'Office Theme')
  t.is(theme.colors.accent1, '4F81BD')
  t.is(theme.colors.dark1, '000000', 'a system colour reads as its last value')
  t.alike(theme.fonts, { heading: 'Calibri', body: 'Calibri' })
  t.alike(pptx.create({ fonts: { heading: 'Georgia' } }).theme().fonts, {
    heading: 'Georgia',
    body: 'Calibri'
  })
})

test('setTheme edits colours and fonts in the master theme only', (t) => {
  const deck = open()
  const before = unzip(deck.save())
  t.alike(
    deck.setTheme({ colors: { accent1: '#1a73e8', dark1: '202124' }, fonts: { body: 'Inter' } }),
    {
      changed: ['accent1', 'dark1', 'body'],
      parts: 1
    }
  )
  const again = reopen(deck).theme()
  t.is(again.colors.accent1, '1A73E8')
  t.is(again.colors.dark1, '202124')
  t.alike(again.fonts, { heading: 'Calibri', body: 'Inter' })
  const after = unzip(deck.save())
  t.ok(decode(after['ppt/theme/theme1.xml']).includes('<a:dk1><a:srgbClr val="202124"/></a:dk1>'))
  t.alike(
    after['ppt/theme/theme2.xml'],
    before['ppt/theme/theme2.xml'],
    'the notes theme is untouched'
  )
})

test('setTheme refuses unknown keys and bad values', (t) => {
  const deck = open()
  t.is(
    codeOf(() => deck.setTheme(JSON.parse('{"colors":{"accent9":"FFFFFF"}}'))),
    'INVALID_OPTIONS'
  )
  t.is(
    codeOf(() => deck.setTheme({ colors: { accent1: 'blue' } })),
    'INVALID_OPTIONS'
  )
  t.is(
    codeOf(() => deck.setTheme({ fonts: { body: ' ' } })),
    'INVALID_OPTIONS'
  )
  t.is(
    codeOf(() => deck.setTheme({})),
    'INVALID_OPTIONS'
  )
})

test('setLayout carries placeholders by role and adds the new ones', (t) => {
  const deck = open()
  t.alike(deck.setLayout(1, 'Two Content'), {
    layout: 'Two Content',
    carried: [
      { id: 2, from: 'title', to: 'title' },
      { id: 3, from: 'obj', to: 'obj' }
    ],
    kept: [],
    removed: [],
    added: [4]
  })
  const slide = reopen(deck).read()[1]
  t.is(slide.layout, 'Two Content')
  t.is(slide.shapes[1].paragraphs?.[0].text, 'Sunlight holds every colour')
  t.is(slide.shapes[1].box?.width, 4038600, 'the body takes the left column of the new layout')
  t.is(slide.notes, 'Mention the sunset question.')
})

test('setLayout keeps text it has no place for, where it was', (t) => {
  const deck = open()
  const subtitle = deck.read()[0].shapes[1].box
  t.alike(deck.setLayout(0, 'Title Only'), {
    layout: 'Title Only',
    carried: [{ id: 2, from: 'ctrTitle', to: 'title' }],
    kept: [{ id: 3, placeholder: 'subTitle', text: 'A short tour of Rayleigh scattering' }],
    removed: [],
    added: []
  })
  const slide = reopen(deck).read()[0]
  t.is(slide.shapes[0].placeholder, 'title')
  t.alike(slide.shapes[1].box, subtitle)
})

test('setLayout drops empty placeholders it cannot carry', (t) => {
  const deck = pptx.create()
  deck.addSlide({ layout: 'Two Content', title: 'Compare' })
  const result = deck.setLayout(0, 'Title Only')
  t.alike(result.removed, [3, 4])
  t.alike(
    reopen(deck)
      .read()[0]
      .shapes.map((s) => s.placeholder),
    ['title']
  )
  t.is(
    codeOf(() => deck.setLayout(0, 'Nope')),
    'LAYOUT_NOT_FOUND'
  )
})
