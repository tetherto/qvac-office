import test from 'brittle'
import * as pptx from '../lib/pptx/index.ts'
import { Package } from '../lib/opc/package.ts'
import * as units from '../lib/units.ts'
import * as xml from '../lib/xml.ts'
import { fixture, decode, unzip } from './helpers/engine.ts'
const reopen = (deck: pptx.Presentation): pptx.Presentation => pptx.open(deck.save())

function bad<T>(value: unknown): T {
  return value as T
}

test('create builds an empty deck with the standard layouts', (t) => {
  const deck = pptx.create()
  t.is(deck.slideCount, 0)
  t.alike(deck.slideSize, { width: 12192000, height: 6858000 })
  t.alike(
    deck.layouts().map((l) => l.name),
    ['Title Slide', 'Title and Content', 'Section Header', 'Two Content', 'Title Only', 'Blank']
  )
  t.ok(deck.changed)

  const again = reopen(deck)
  t.is(again.slideCount, 0)
  t.is(again.layouts().length, 6)
})

test('create writes a complete package', (t) => {
  const files = unzip(pptx.create({ title: 'Q3 Review' }).save())
  for (const name of [
    '[Content_Types].xml',
    '_rels/.rels',
    'ppt/presentation.xml',
    'ppt/_rels/presentation.xml.rels',
    'ppt/slideMasters/slideMaster1.xml',
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    'ppt/slideLayouts/slideLayout1.xml',
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    'ppt/theme/theme1.xml',
    'ppt/presProps.xml',
    'ppt/viewProps.xml',
    'ppt/tableStyles.xml',
    'docProps/core.xml',
    'docProps/app.xml'
  ]) {
    t.ok(files[name], name)
  }

  t.ok(decode(files['docProps/core.xml']).includes('<dc:title>Q3 Review</dc:title>'))
  t.is(Object.keys(files)[0], '[Content_Types].xml')
})

test('create takes a size, theme colours and fonts', (t) => {
  const deck = pptx.create({
    size: '4:3',
    colors: { accent1: '#C0392B' },
    fonts: { heading: 'Georgia', body: 'Verdana' }
  })
  t.alike(deck.slideSize, { width: 9144000, height: 6858000 })
  const theme = decode(unzip(deck.save())['ppt/theme/theme1.xml'])
  t.ok(theme.includes('<a:accent1><a:srgbClr val="C0392B"/></a:accent1>'))
  t.ok(theme.includes('<a:majorFont><a:latin typeface="Georgia"/>'))
  t.ok(theme.includes('<a:minorFont><a:latin typeface="Verdana"/>'))

  t.alike(pptx.create({ size: { width: 9144000, height: 5143500 } }).slideSize, {
    width: 9144000,
    height: 5143500
  })
  t.exception(() => pptx.create({ size: bad('21:9') }), /size must be/)
  t.exception(
    () => pptx.create({ colors: bad({ accent9: '000000' }) }),
    /unknown theme color "accent9"/
  )
  t.exception(() => pptx.create({ colors: { accent1: 'red' } }), /must be hex RGB/)
  t.exception(() => pptx.create({ fonts: bad({ title: 'Arial' }) }), /unknown theme font "title"/)
})

test('a created deck takes slides on every layout', (t) => {
  const deck = pptx.create()
  deck.addSlide({
    layout: 'Title Slide',
    title: 'Why the Sky Is Blue',
    body: 'Rayleigh scattering'
  })
  deck.addSlide({
    title: 'Scattering',
    body: ['Short waves scatter', { text: 'Blue most', level: 1 }]
  })
  deck.addSlide({ layout: 'Section Header', title: 'Evidence', body: 'What we measured' })
  const two = deck.addSlide({ layout: 'Two Content', title: 'Compare' })
  deck.setText({ slide: two.index, shape: two.shapes[1].id! }, 'Left side')
  deck.setText({ slide: two.index, shape: two.shapes[2].id! }, 'Right side')
  deck.addSlide({ layout: 'Title Only', title: 'Chart' })
  deck.addSlide({ layout: 'Blank' })

  const slides = reopen(deck).read()
  t.alike(
    slides.map((s) => s.layout),
    ['Title Slide', 'Title and Content', 'Section Header', 'Two Content', 'Title Only', 'Blank']
  )
  t.alike(
    slides[0].shapes.map((s) => s.placeholder),
    ['ctrTitle', 'subTitle']
  )
  t.alike(slides[1].shapes[1].paragraphs, [
    { text: 'Short waves scatter', level: 0 },
    { text: 'Blue most', level: 1 }
  ])
  t.alike(
    slides[3].shapes.map((s) => s.paragraphs![0].text),
    ['Compare', 'Left side', 'Right side']
  )
  t.is(slides[5].shapes.length, 0)
})

test('the default layout of a new deck is Title and Content', (t) => {
  const deck = pptx.create()
  t.is(deck.addSlide().layout, 'Title and Content')
})

test('addPicture with no position fills an empty content placeholder in its place', (t) => {
  const deck = pptx.create()
  deck.addSlide({ layout: 'Two Content', title: 'Imaging', body: ['Scans read faster'] })
  const right = deck.read()[0].shapes[2].box
  const placed = deck.addPicture(0, fixture('red.png'))
  const shapes = reopen(deck).read()[0].shapes

  t.alike(
    shapes.map((s) => s.kind),
    ['text', 'text', 'picture'],
    'the empty placeholder is replaced, not covered'
  )
  t.ok(right && placed.x >= right.x && placed.x + placed.width <= right.x + right.width)
  t.ok(right && placed.y >= right.y && placed.y + placed.height <= right.y + right.height)
})

test('addPicture with no position on a title slide sits below the text, not over it', (t) => {
  const deck = pptx.create()
  deck.addSlide({ layout: 'Title Slide', title: 'AI in Healthcare', body: 'Where it helps' })
  const text = deck.read()[0].shapes.map((s) => s.box)
  const placed = deck.addPicture(0, fixture('red.png'))
  const bottom = Math.max(...text.map((box) => (box ? box.y + box.height : 0)))

  t.ok(placed.y >= bottom, 'the picture starts below the subtitle')
  t.ok(placed.y + placed.height <= deck.slideSize.height)
})

test('addPicture on a slide full of text splits it, text on the left and picture on the right', (t) => {
  const deck = pptx.create()
  deck.addSlide({
    title: 'Fire',
    body: ['Lightning lit the first fires', 'People learned to keep them']
  })
  const before = deck.read()[0].shapes[1].box
  const placed = deck.addPicture(0, fixture('red.png'))
  const body = reopen(deck).read()[0].shapes[1].box

  t.ok(before && body && body.width < before.width / 2 + 1, 'the text box takes the left half')
  t.ok(body && placed.x >= body.x + body.width, 'the picture starts right of the text')
  t.ok(before && placed.x + placed.width <= before.x + before.width)
})

test('addPicture on a cover slide with no subtitle sits below the title, not in the subtitle box', (t) => {
  const deck = pptx.create()
  deck.addSlide({ layout: 'Title Slide', title: 'Fire' })
  const [title, subtitle] = deck.read()[0].shapes.map((s) => s.box)
  const placed = deck.addPicture(0, photo(1600, 1200))

  t.alike(
    reopen(deck)
      .read()[0]
      .shapes.map((s) => s.kind),
    ['text', 'picture'],
    'the empty subtitle under the picture goes'
  )
  t.ok(title && placed.y >= title.y + title.height)
  t.ok(subtitle && placed.height > subtitle.height, 'the picture is not squeezed into the subtitle')
})

test('a second picture on a slide of bullets shares the half the first one took, not the text', (t) => {
  const deck = pptx.create()
  deck.addSlide({
    title: 'Fire',
    body: ['Lightning lit the first fires', 'People learned to keep them']
  })
  deck.addPicture(0, photo(1600, 1200))
  const body = deck.read()[0].shapes[1].box
  const second = deck.addPicture(0, photo(1600, 1200))
  const [, text, first] = reopen(deck)
    .read()[0]
    .shapes.map((s) => s.box)

  t.alike(text, body, 'the text keeps its half')
  t.ok(body && first && first.x >= body.x + body.width, 'the first picture stays right of the text')
  t.ok(first && first.x + first.width <= second.x, 'the pictures sit side by side')
})

test('addPicture with a size but no position shrinks to the half beside the text', (t) => {
  const deck = pptx.create()
  deck.addSlide({ title: 'Fire', body: ['Lightning lit the first fires'] })
  const wide = deck.addPicture(0, photo(1600, 900), { width: units.inches(8) })
  const tall = deck.addPicture(0, photo(900, 1600), { width: units.inches(6) })
  const [title, body] = deck.read()[0].shapes.map((s) => s.box)
  const { height } = deck.slideSize

  for (const placed of [wide, tall]) {
    t.ok(body && placed.x >= body.x + body.width, 'the picture stays right of the text')
    t.ok(title && placed.y >= title.y + title.height, 'the picture stays below the title')
    t.ok(placed.y + placed.height <= height, 'the picture stays on the slide')
  }

  t.is(Math.round((wide.height / wide.width) * 16), 9, 'the picture keeps its shape')
})

test('addSlide takes an image and places it like addPicture', (t) => {
  const deck = pptx.create()
  deck.addSlide({ title: 'Only a picture', image: fixture('red.png') })
  const added = deck.addSlide({ title: 'Both', body: ['A point'], image: fixture('red.png') })

  t.alike(
    deck.read()[0].shapes.map((s) => s.kind),
    ['text', 'picture'],
    'the empty body becomes the picture'
  )
  t.alike(
    added.shapes.map((s) => s.kind),
    ['text', 'text', 'picture']
  )
  t.exception(() => deck.addSlide({ title: 'Bad', image: new Uint8Array(40) }), /only PNG, JPEG/)
  t.is(deck.slideCount, 2, 'a bad image adds no slide')
})

test('addPicture with a position keeps it, even over an empty placeholder', (t) => {
  const deck = pptx.create()
  deck.addSlide({ title: 'Placed' })
  const placed = deck.addPicture(0, fixture('red.png'), { x: units.inches(1), y: units.inches(1) })

  t.is(placed.x, units.inches(1))
  t.is(placed.y, units.inches(1))
  t.is(deck.read()[0].shapes.length, 3, 'the placeholder stays')
})

test('addTextBox places styled text anywhere', (t) => {
  const deck = pptx.create()
  deck.addSlide({ layout: 'Blank' })
  const box = deck.addTextBox(0, {
    text: ['Hand placed', 'Second line'],
    x: units.inches(1),
    y: units.inches(1),
    width: units.inches(5),
    height: units.inches(1),
    style: { size: 28, bold: true, color: '1A73E8', align: 'center' },
    fill: 'F1F3F4'
  })
  t.alike(box, { id: 2, x: 914400, y: 914400, width: 4572000, height: 914400, overlaps: [] })
  const read = reopen(deck).read()[0].shapes[0]
  t.is(read.kind, 'text')
  t.alike(
    read.paragraphs!.map((p) => p.text),
    ['Hand placed', 'Second line']
  )
  const xml = decode(unzip(deck.save())['ppt/slides/slide1.xml'])
  t.ok(xml.includes('<p:cNvSpPr txBox="1"/>'))
  t.ok(xml.includes('<a:srgbClr val="F1F3F4"/>'))
  t.ok(xml.includes('<a:pPr algn="ctr"/>'))
  t.ok(xml.includes('<a:rPr sz="2800" b="1"><a:solidFill><a:srgbClr val="1A73E8"/>'))
  t.exception(() => deck.addTextBox(0, { text: 'x', fill: 'grey' }), /fill must be a hex RGB/)
  t.exception(() => deck.addTextBox(0, { text: 'x', width: -1 }), /non-negative whole number/)
})

test('addTable builds a styled table', (t) => {
  const deck = pptx.create()
  deck.addSlide({ layout: 'Title Only', title: 'Numbers' })
  const placed = deck.addTable(0, {
    rows: [
      ['Colour', 'Wavelength', 'Share'],
      ['Blue', '450 nm', 0.6],
      ['Red', '700 nm']
    ],
    width: units.inches(9),
    height: units.inches(1.5)
  })
  const read = reopen(deck)
    .read()[0]
    .shapes.find((s) => s.id === placed.id)!
  t.is(read.kind, 'table')
  t.alike(read.rows, [
    ['Colour', 'Wavelength', 'Share'],
    ['Blue', '450 nm', '0.6'],
    ['Red', '700 nm', '']
  ])
  const xml = decode(unzip(deck.save())['ppt/slides/slide1.xml'])
  t.ok(xml.includes('<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>'))
  t.ok(xml.includes(`<a:gridCol w="${Math.floor(units.inches(9) / 3)}"/>`))
  deck.setCellText({ slide: 0, shape: placed.id, row: 2, column: 2 }, '0.4')
  t.is(
    reopen(deck)
      .read()[0]
      .shapes.find((s) => s.id === placed.id)!.rows![2][2],
    '0.4'
  )
  t.exception(() => deck.addTable(0, { rows: [] }), /non-empty array of rows/)
  t.exception(() => deck.addTable(0, { rows: bad([[{}]]) }), /cell 0,0 must be/)
})

test('setBackground sets and clears a slide background', (t) => {
  const deck = pptx.open(fixture('basic.pptx'))
  deck.setBackground(0, '#101820')
  let xml = decode(unzip(deck.save())['ppt/slides/slide1.xml'])
  t.ok(
    xml.includes(
      '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="101820"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
    )
  )
  deck.setBackground(0, null)
  xml = decode(unzip(deck.save())['ppt/slides/slide1.xml'])
  t.absent(xml.includes('<p:bg>'))
  t.exception(() => deck.setBackground(0, 'navy'), /hex RGB/)
})

test('setStyle aligns paragraphs without touching runs', (t) => {
  const deck = pptx.open(fixture('basic.pptx'))
  deck.setStyle({ slide: 0, shape: 'title' }, { align: 'right' })
  const xml = decode(unzip(deck.save())['ppt/slides/slide1.xml'])
  t.ok(xml.includes('<a:p><a:pPr algn="r"/><a:r><a:t>Why the Sky Is Blue</a:t></a:r></a:p>'))
  t.exception(
    () => deck.setStyle({ slide: 0, shape: 'title' }, bad({ align: 'middle' })),
    /align must be/
  )
})

test('a deck built from scratch survives a full edit cycle', (t) => {
  let deck = pptx.create({ title: 'Cycle' })
  deck.addSlide({ layout: 'Title Slide', title: 'Cover' })
  deck.addSlide({ title: 'Body', body: 'Text' })
  deck.addChart(1, {
    type: 'line',
    categories: ['a', 'b'],
    series: [{ name: 's', values: [1, 2] }]
  })
  deck.addPicture(0, fixture('red.png'))
  deck = reopen(deck)
  deck.setChartData(
    { slide: 1, shape: 4 },
    {
      categories: ['a', 'b', 'c'],
      series: [{ name: 's', values: [1, 2, 3] }]
    }
  )
  deck.moveSlide(1, { to: 0 })
  deck = reopen(deck)
  t.alike(
    deck.read().map((s) => s.shapes[0].paragraphs![0].text),
    ['Body', 'Cover']
  )
  t.alike(deck.read()[0].shapes[2].chart!.series[0].values, [1, 2, 3])
})

function photo(width: number, height: number): Uint8Array {
  const bytes = fixture('red.png').slice()
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  // placement reads only the PNG header's size
  view.setUint32(16, width)
  view.setUint32(20, height)

  return bytes
}
