import test from 'brittle'
import * as pptx from '../lib/pptx/index.ts'
import { Package } from '../lib/opc/package.ts'
import * as units from '../lib/units.ts'
import * as xml from '../lib/xml.ts'
import { fixture, decode, unzip } from './helpers/engine.ts'
const open = (): pptx.Presentation => pptx.open(fixture('basic.pptx'))
const reopen = (deck: pptx.Presentation): pptx.Presentation => pptx.open(deck.save())

function shape(deck: pptx.Presentation, slide: number, id: number): pptx.ShapeInfo {
  const found = deck.read()[slide].shapes.find((s) => s.id === id)
  if (!found) {
    throw new Error(`no shape ${id}`)
  }

  return found
}

function bad<T>(value: unknown): T {
  return value as T
}
function slidePart(deck: pptx.Presentation, slide: number): string {
  const list = xml.child(deck.package.root(deck.main), xml.ns.p, 'sldIdLst')
  const ids = list ? xml.children(list, xml.ns.p, 'sldId') : []
  const rel = deck.package.rels(deck.main).get(ids[slide]?.getAttributeNS(xml.ns.r, 'id') ?? null)
  return rel?.part ?? ''
}

const slideXml = (deck: pptx.Presentation, slide: number): string =>
  decode(deck.package.bytes(slidePart(deck, slide)))
function codeOf(fn: () => unknown): string | null {
  try {
    fn()
  } catch (err) {
    return err instanceof Error && 'code' in err ? String(err.code) : null
  }

  return null
}

test('reads slides, layouts, shapes, tables, groups and notes', (t) => {
  const slides = open().read()
  t.is(slides.length, 3)
  t.alike(
    slides.map((s) => s.layout),
    ['Title Slide', 'Title and Content', 'Title Only']
  )
  t.alike(slides[1].shapes[1].paragraphs, [
    { text: 'Sunlight holds every colour', level: 0 },
    { text: 'Short wavelengths scatter more', level: 1 }
  ])
  t.is(slides[1].notes, 'Mention the sunset question.')
  const kinds = slides[2].shapes.map((s) => s.kind)
  t.alike(kinds, ['text', 'picture', 'table', 'group', 'text'])
  t.alike(slides[2].shapes[2].rows, [
    ['Colour', 'Wavelength'],
    ['Blue', '450 nm']
  ])
  t.is(slides[2].shapes[4].group, 5)
})

test('rejects files that are not presentations', (t) => {
  const pkg = Package.open(fixture('basic.pptx'))
  pkg.contentTypes.setOverride('/ppt/presentation.xml', 'application/xml')
  t.is(
    codeOf(() => new pptx.Presentation(pkg)),
    'NOT_PPTX'
  )
})

test('setText keeps run and paragraph formatting', (t) => {
  const deck = open()
  const result = deck.setText({ slide: 1, shape: 3 }, [
    'Light is a mix of colours',
    'Blue scatters about five times more than red'
  ])
  t.alike(result, { paragraphs: 2, changed: true, overflow: false })

  const again = reopen(deck)
  t.alike(shape(again, 1, 3).paragraphs, [
    { text: 'Light is a mix of colours', level: 0 },
    { text: 'Blue scatters about five times more than red', level: 1 }
  ])
  const xml = slideXml(again, 1)
  t.ok(xml.includes('<a:rPr b="1" sz="2400"/><a:t>Light is a mix of colours</a:t>'))
  t.ok(xml.includes('<a:pPr lvl="1"/><a:r><a:t>Blue scatters'))
})

test('setText splits a string on newlines and extends with the last paragraph style', (t) => {
  const deck = open()
  deck.setText({ slide: 1, shape: 'body' }, 'one\ntwo\nthree')
  const paragraphs = shape(reopen(deck), 1, 3).paragraphs ?? []
  t.alike(
    paragraphs.map((p) => p.text),
    ['one', 'two', 'three']
  )
  t.alike(
    paragraphs.map((p) => p.level),
    [0, 1, 1]
  )
})

test('setText sets explicit levels and line breaks', (t) => {
  const deck = open()
  deck.setText({ slide: 1, shape: 'body' }, [
    { text: 'top', level: 0 },
    { text: 'first\nsecond', level: 2 }
  ])
  const again = reopen(deck)
  t.alike(shape(again, 1, 3).paragraphs, [
    { text: 'top', level: 0 },
    { text: 'first\nsecond', level: 2 }
  ])
  t.ok(slideXml(again, 1).includes('<a:br/>'))
})

test('addSlide whose layout cannot hold its text adds no slide', (t) => {
  const deck = pptx.create()
  deck.addSlide({ layout: 'Title Slide', title: 'Cover' })
  t.exception(() => deck.addSlide({ layout: 'Title Only', title: 'Reason', body: ['Flaky'] }))
  t.is(deck.slideCount, 1)
})

test('setText with the current text is not a change', (t) => {
  const deck = open()
  t.alike(deck.setText({ slide: 0, shape: 'title' }, 'Why the Sky Is Blue'), {
    paragraphs: 1,
    changed: false,
    overflow: false
  })
  t.is(deck.changed, false)
})

test('setText targets placeholders by role', (t) => {
  const deck = open()
  deck.setText({ slide: 0, shape: 'title' }, 'Why Sunsets Are Red')
  deck.setText({ slide: 0, shape: 'subtitle' }, 'And why the sky is blue')
  deck.setText({ slide: 2, shape: 6 }, 'A caption inside a group')
  const slides = reopen(deck).read()
  t.is(slides[0].shapes[0].paragraphs![0].text, 'Why Sunsets Are Red')
  t.is(slides[0].shapes[1].paragraphs![0].text, 'And why the sky is blue')
  t.is(slides[2].shapes[4].paragraphs![0].text, 'A caption inside a group')
})

test('errors name what exists', (t) => {
  const deck = open()
  t.exception(
    () => deck.setText({ slide: 7, shape: 'title' }, 'x'),
    /slide 7 does not exist; the deck has 3 slides \(0 to 2\)/
  )
  t.exception(
    () => deck.setText({ slide: 0, shape: 99 }, 'x'),
    /slide 0 has no shape 99; its shapes are: 2 text \(ctrTitle\) "Title 1"/
  )
  t.is(
    codeOf(() => deck.setText({ slide: 2, shape: 3 }, 'x')),
    'NOT_TEXT'
  )
  t.exception(
    () => deck.setText({ slide: 2, shape: 4 }, 'x'),
    /is a table; set its cells with setCellText/
  )
  t.is(
    codeOf(() => deck.setText({ slide: 0, shape: 'title' }, [{ text: 'x', level: 12 }])),
    'INVALID_TEXT'
  )
  t.is(deck.changed, false)
})

test('setCellText writes one table cell', (t) => {
  const deck = open()
  deck.setCellText({ slide: 2, shape: 4, row: 1, column: 1 }, '~450 nm')
  t.alike(shape(reopen(deck), 2, 4).rows, [
    ['Colour', 'Wavelength'],
    ['Blue', '~450 nm']
  ])
  t.exception(
    () => deck.setCellText({ slide: 2, shape: 4, row: 5, column: 0 }, 'x'),
    /row 5 does not exist; the table has 2/
  )
  t.exception(() => deck.setCellText({ slide: 2, shape: 2, row: 0, column: 0 }, 'x'), /not a table/)
})

test('setStyle sets size, weight, colour and font on every run', (t) => {
  const deck = open()
  const { runs } = deck.setStyle(
    { slide: 0, shape: 'title' },
    {
      size: 44,
      bold: true,
      color: '#1a73e8',
      font: 'Georgia'
    }
  )
  t.is(runs, 1)
  const xml = slideXml(reopen(deck), 0)
  t.ok(
    xml.includes(
      '<a:rPr sz="4400" b="1"><a:solidFill><a:srgbClr val="1A73E8"/></a:solidFill><a:latin typeface="Georgia"/></a:rPr>'
    )
  )
})

test('setStyle replaces an existing fill in schema order', (t) => {
  const deck = open()
  deck.setStyle({ slide: 1, shape: 'body' }, { color: 'FF0000' })
  deck.setStyle({ slide: 1, shape: 'body' }, { font: 'Arial', color: '00FF00' })
  const xml = slideXml(deck, 1)
  t.absent(xml.includes('FF0000'))
  t.ok(
    xml.includes('<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:latin typeface="Arial"/>')
  )
})

test('setStyle can target one paragraph or a whole table', (t) => {
  const deck = open()
  deck.setStyle({ slide: 1, shape: 'body' }, { italic: true }, { paragraph: 1 })
  const xml = slideXml(deck, 1)
  t.ok(xml.includes('<a:rPr b="1" sz="2400"/><a:t>Sunlight'))
  t.ok(xml.includes('<a:rPr i="1"/><a:t>Short wavelengths'))
  t.is(deck.setStyle({ slide: 2, shape: 4 }, { size: 14 }).runs, 4)
  t.exception(
    () => deck.setStyle({ slide: 1, shape: 'body' }, { italic: true }, { paragraph: 4 }),
    /paragraph 4 does not exist/
  )
})

test('setStyle validates its input', (t) => {
  const deck = open()
  t.exception(() => deck.setStyle({ slide: 0, shape: 'title' }, {}), /names nothing to change/)
  t.exception(() => deck.setStyle({ slide: 0, shape: 'title' }, { size: 0 }), /size is in points/)
  t.exception(
    () => deck.setStyle({ slide: 0, shape: 'title' }, bad({ colour: 'red' })),
    /unknown style "colour"/
  )
  t.exception(() => deck.setStyle({ slide: 0, shape: 'title' }, { color: 'blue' }), /hex RGB/)
  t.exception(() => deck.setStyle({ slide: 2, shape: 3 }, { bold: true }), /holds no text/)
  t.is(deck.changed, false)
})

test('addSlide reuses a layout with a body, never the cover', (t) => {
  const deck = open()
  const added = deck.addSlide({
    title: 'What Changed',
    body: ['One new slide', 'Nothing else touched']
  })
  t.is(added.index, 3)
  t.is(added.layout, 'Title and Content')

  const coverOnly = open()
  coverOnly.deleteSlide(2)
  coverOnly.deleteSlide(1)
  t.is(coverOnly.addSlide().layout, 'Title and Content')

  const again = reopen(deck)
  const slide = again.read()[3]
  t.is(slide.layout, 'Title and Content')
  t.is(slide.shapes[0].paragraphs![0].text, 'What Changed')
})

test('addSlide writes the part, relationships, content type and slide id', (t) => {
  const deck = open()
  const added = deck.addSlide({
    layout: 'title and content',
    title: 'Summary',
    body: 'Short light scatters most'
  })
  t.is(added.layout, 'Title and Content')
  const files = unzip(deck.save())
  t.ok(files['ppt/slides/slide4.xml'])
  t.ok(
    decode(files['ppt/slides/_rels/slide4.xml.rels']).includes(
      'Target="../slideLayouts/slideLayout2.xml"'
    )
  )
  t.ok(decode(files['[Content_Types].xml']).includes('PartName="/ppt/slides/slide4.xml"'))
  t.ok(decode(files['ppt/_rels/presentation.xml.rels']).includes('Id="rId11"'))
  t.ok(decode(files['ppt/presentation.xml']).includes('<p:sldId id="259" r:id="rId11"/>'))

  const read = pptx.open(deck.save()).read()[3]
  t.alike(
    read.shapes.map((s) => s.placeholder),
    ['title', 'obj']
  )
  t.is(read.shapes[1].paragraphs![0].text, 'Short light scatters most')
})

test('addSlide inserts at a position and copies a slide layout', (t) => {
  const deck = open()
  deck.addSlide({ like: 0, at: 1, title: 'Second cover' })
  const slides = reopen(deck).read()
  t.alike(
    slides.map((s) => s.layout),
    ['Title Slide', 'Title Slide', 'Title and Content', 'Title Only']
  )
  t.is(slides[1].shapes[0].paragraphs![0].text, 'Second cover')
  t.exception(
    () => deck.addSlide({ layout: 'Nope' }),
    /no layout is named "Nope"; the deck has "Title Slide"/
  )
  t.exception(() => deck.addSlide({ at: 9 }), /at must be a position from 0 to 4/)
  t.exception(() => deck.addSlide({ like: 0, layout: 'Blank' }), /either like or layout/)
})

test('setText turns a vertical tab into a line break and drops other control characters', (t) => {
  const deck = open()
  deck.setText({ slide: 0, shape: 'title' }, 'a\u000bb\u0001c')
  const slide = slideXml(deck, 0)
  t.absent(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(slide))
  t.ok(slide.includes('<a:t>a</a:t></a:r><a:br'))
  t.ok(slide.includes('<a:t>bc</a:t>'))
  t.is(reopen(deck).read()[0].shapes[0].paragraphs![0].text, 'a\nbc')
})

test('a failed addSlide leaves the deck as it was', (t) => {
  const deck = open()
  t.is(
    codeOf(() => deck.addSlide({ layout: 'Blank', title: 'Agenda' })),
    'SHAPE_NOT_FOUND'
  )
  t.is(
    codeOf(() => deck.addSlide({ title: bad(42) })),
    'INVALID_TEXT'
  )
  t.is(deck.changed, false)
  t.is(reopen(deck).read().length, 3)
  t.absent(unzip(deck.save())['ppt/slides/slide4.xml'])
})

test('addPicture embeds media sized by aspect ratio', (t) => {
  const deck = open()
  const red = fixture('red.png')
  const placed = deck.addPicture(1, red, {
    x: units.inches(1),
    y: units.inches(1),
    width: units.inches(4)
  })
  t.alike(placed, { id: 4, x: 914400, y: 914400, width: 3657600, height: 1828800 })

  const files = unzip(deck.save())
  t.ok(files['ppt/media/image2.png'])
  t.ok(decode(files['ppt/slides/_rels/slide2.xml.rels']).includes('Target="../media/image2.png"'))
  const again = pptx.open(deck.save())
  t.alike(
    again.read()[1].shapes.map((s) => s.kind),
    ['text', 'text', 'picture']
  )
})

test('addPicture centres a picture at its native size and adds a content type', (t) => {
  const deck = open()
  const gif = new Uint8Array(32)
  gif.set([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 10, 0, 5, 0])
  const placed = deck.addPicture(0, gif)
  t.is(placed.width, 10 * 9525)
  t.is(placed.height, 5 * 9525)
  t.is(placed.x, Math.round((9144000 - placed.width) / 2))
  const types = decode(unzip(deck.save())['[Content_Types].xml'])
  t.ok(types.includes('<Default Extension="gif" ContentType="image/gif"/>'))
  t.exception(() => deck.addPicture(0, new Uint8Array(40)), /only PNG, JPEG and GIF/)
  t.exception(() => deck.addPicture(0, fixture('red.png'), { width: 1.5 }), /whole number of EMU/)
})

test('deleteSlide removes the slide, its notes and orphaned parts only', (t) => {
  const deck = open()
  deck.deleteSlide(1)
  t.is(deck.slideCount, 2)
  const files = unzip(deck.save())
  t.absent(files['ppt/slides/slide2.xml'])
  t.absent(files['ppt/notesSlides/notesSlide1.xml'])
  t.ok(files['ppt/notesMasters/notesMaster1.xml'])
  t.ok(files['ppt/media/image1.png'])
  t.absent(decode(files['[Content_Types].xml']).includes('/ppt/slides/slide2.xml'))
  t.alike(
    pptx
      .open(deck.save())
      .read()
      .map((s) => s.layout),
    ['Title Slide', 'Title Only']
  )
})

test('deleteSlide drops media nothing else uses', (t) => {
  const deck = open()
  deck.deleteSlide(2)
  t.absent(unzip(deck.save())['ppt/media/image1.png'])
})

test('deleteSlide refuses a slide another slide links to', (t) => {
  const deck = open()
  deck.package
    .rels('/ppt/slides/slide1.xml')
    .add(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
      '/ppt/slides/slide3.xml'
    )
  t.exception(
    () => deck.deleteSlide(2),
    /slide 2 is the target of a link on slide 0; remove the link first/
  )
})

test('moveSlide reorders slides', (t) => {
  const deck = open()
  t.alike(deck.moveSlide(0, { to: 2 }), { moved: true })
  t.alike(deck.moveSlide(1, { to: 1 }), { moved: false })
  t.alike(
    reopen(deck)
      .read()
      .map((s) => s.layout),
    ['Title and Content', 'Title Only', 'Title Slide']
  )
  t.exception(() => deck.moveSlide(0, { to: 3 }), /to must be a position from 0 to 2/)
})

test('slide edits keep sections consistent', (t) => {
  const deck = withSections(open(), [['256'], ['257', '258']])
  deck.moveSlide(2, { to: 0 })
  t.alike(sections(deck), [['258', '256'], ['257']])

  deck.addSlide({ at: 3 })
  t.alike(sections(deck), [
    ['258', '256'],
    ['257', '259']
  ])

  deck.deleteSlide(1)
  t.alike(sections(deck), [['258'], ['257', '259']])

  const again = reopen(deck)
  t.alike(sections(again), [['258'], ['257', '259']])
})

test('an edited deck survives repeated open and save', (t) => {
  let deck = open()
  deck.addSlide({ title: 'A', body: 'B' })
  deck.addPicture(3, fixture('red.png'))
  deck.moveSlide(3, { to: 1 })
  deck = reopen(deck)
  deck.deleteSlide(0)
  deck = reopen(deck)
  t.is(deck.slideCount, 3)
  t.alike(
    deck.read().map((s) => s.shapes[0].paragraphs![0].text),
    ['A', 'Scattering', 'Measurements']
  )
})

function withSections(deck: pptx.Presentation, groups: string[][]): pptx.Presentation {
  const doc = deck.package.xml(deck.main)
  const P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
  const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
  const extLst = doc.createElementNS(P, 'p:extLst')
  const ext = doc.createElementNS(P, 'p:ext')
  ext.setAttribute('uri', '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}')
  const list = doc.createElementNS(P14, 'p14:sectionLst')
  groups.forEach((ids, i) => {
    const section = doc.createElementNS(P14, 'p14:section')
    section.setAttribute('name', `Section ${i}`)
    section.setAttribute('id', `{00000000-0000-0000-0000-00000000000${i}}`)
    const sldIdLst = doc.createElementNS(P14, 'p14:sldIdLst')
    for (const id of ids) {
      const ref = doc.createElementNS(P14, 'p14:sldId')
      ref.setAttribute('id', id)
      sldIdLst.appendChild(ref)
    }

    section.appendChild(sldIdLst)
    list.appendChild(section)
  })
  ext.appendChild(list)
  extLst.appendChild(ext)
  xml.root(doc).appendChild(extLst)
  deck.package.touch(deck.main)

  return reopen(deck)
}

function sections(deck: pptx.Presentation): string[][] {
  const doc = deck.package.xml(deck.main)
  const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
  return Array.from(doc.getElementsByTagNameNS(P14, 'sldIdLst')).map((list) =>
    Array.from(list.getElementsByTagNameNS(P14, 'sldId')).map((el) => xml.attr(el, 'id'))
  )
}
