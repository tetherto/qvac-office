import test from 'brittle'
import * as pptx from '../lib/pptx/index.ts'
import { Package } from '../lib/opc/package.ts'
import * as units from '../lib/units.ts'
import * as xml from '../lib/xml.ts'
import * as utf8 from '../lib/utf8.ts'
import { fixture, decode, unzip, zipFiles } from './helpers/engine.ts'
const open = (): pptx.Presentation => pptx.open(fixture('basic.pptx'))
const reopen = (deck: pptx.Presentation): pptx.Presentation => pptx.open(deck.save())
const files = (deck: pptx.Presentation): Record<string, Uint8Array> => unzip(deck.save())

function bad<T>(value: unknown): T {
  return value as T
}

test('read reports each shape box, inherited from the layout when not set', (t) => {
  const slides = open().read()
  t.alike(slides[1].shapes[0].box, { x: 457200, y: 274638, width: 8229600, height: 1143000 })
  t.alike(slides[2].shapes[1].box, { x: 914400, y: 1828800, width: 1828800, height: 914400 })
  t.ok(slides[2].shapes[2].box)
})

test('deleteShape removes a shape and the media only it used', (t) => {
  const deck = open()
  t.alike(deck.deleteShape({ slide: 2, shape: 3 }), { deleted: 3 })
  t.alike(
    reopen(deck)
      .read()[2]
      .shapes.map((s) => s.kind),
    ['text', 'table', 'group', 'text']
  )
  t.absent(files(deck)['ppt/media/image1.png'])

  const group = open()
  group.deleteShape({ slide: 2, shape: 5 })
  t.alike(
    reopen(group)
      .read()[2]
      .shapes.map((s) => s.id),
    [2, 3, 4]
  )
  t.exception(() => deck.deleteShape({ slide: 2, shape: 3 }), /slide 2 has no shape 3/)
})

test('deleteShape drops animations that target the shape', (t) => {
  const deck = open()
  const doc = deck.package.xml('/ppt/slides/slide2.xml')
  const timing = xml.parse(
    '<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tnLst><p:par><p:cTn id="1"><p:childTnLst><p:set><p:cBhvr><p:tgtEl><p:spTgt spid="3"/></p:tgtEl></p:cBhvr></p:set></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
  )
  xml.root(doc).appendChild(doc.importNode(xml.root(timing), true))
  deck.package.touch('/ppt/slides/slide2.xml')
  t.alike(deck.deleteShape({ slide: 1, shape: 3 }), { deleted: 3, removed: { animations: true } })
  t.absent(decode(files(deck)['ppt/slides/slide2.xml']).includes('p:timing'))
})

test('setBox moves and resizes, filling in inherited geometry', (t) => {
  const deck = open()
  t.alike(deck.setBox({ slide: 1, shape: 'title' }, { y: units.inches(0.5) }), {
    id: 2,
    x: 457200,
    y: 457200,
    width: 8229600,
    height: 1143000
  })
  deck.setBox(
    { slide: 2, shape: 4 },
    { x: 0, y: 0, width: units.inches(5), height: units.inches(2) }
  )
  const again = reopen(deck).read()
  t.alike(again[1].shapes[0].box, { x: 457200, y: 457200, width: 8229600, height: 1143000 })
  t.alike(again[2].shapes[2].box, { x: 0, y: 0, width: 4572000, height: 1828800 })
  t.exception(
    () => deck.setBox({ slide: 1, shape: 'title' }, bad({ depth: 1 })),
    /box takes x, y, width and height/
  )
  t.exception(() => deck.setBox({ slide: 1, shape: 'title' }, { width: 0 }), /no visible size/)
})

test('replacePicture swaps the image and fits it in the old box', (t) => {
  const deck = open()
  const square = new Uint8Array(32)
  square.set([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 10, 0, 10, 0])
  const placed = deck.replacePicture({ slide: 2, shape: 3 }, square)
  t.alike(placed, { id: 3, x: 1371600, y: 1828800, width: 914400, height: 914400 })
  const out = files(deck)
  t.ok(out['ppt/media/image1.gif'])
  t.absent(out['ppt/media/image1.png'])

  const stretched = open()
  t.alike(stretched.replacePicture({ slide: 2, shape: 3 }, square, { stretch: true }), {
    id: 3,
    x: 914400,
    y: 1828800,
    width: 1828800,
    height: 914400
  })
  t.exception(() => deck.replacePicture({ slide: 2, shape: 2 }, square), /is a text, not a picture/)
})

test('replacePicture drops the old crop and the old SVG', (t) => {
  const source = unzip(fixture('basic.pptx'))
  const edit = (name: string, from: string, to: string): void => {
    source[name] = utf8.encode(decode(source[name]).replace(from, to))
  }
  edit(
    'ppt/slides/slide3.xml',
    '<a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch>',
    '<a:blip r:embed="rId2"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId9"/></a:ext></a:extLst></a:blip><a:srcRect l="25000" r="25000"/><a:stretch><a:fillRect t="10000" b="10000"/></a:stretch>'
  )
  edit(
    'ppt/slides/_rels/slide3.xml.rels',
    '</Relationships>',
    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image9.svg"/></Relationships>'
  )
  edit(
    '[Content_Types].xml',
    '<Default ',
    '<Default Extension="svg" ContentType="image/svg+xml"/><Default '
  )
  source['ppt/media/image9.svg'] = utf8.encode('<svg xmlns="http://www.w3.org/2000/svg"/>')
  const deck = pptx.open(zipFiles(source))
  const square = new Uint8Array(32)
  square.set([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 10, 0, 10, 0])
  deck.replacePicture({ slide: 2, shape: 3 }, square)
  const out = files(deck)
  const slide = decode(out['ppt/slides/slide3.xml'])
  t.absent(slide.includes('srcRect'))
  t.absent(slide.includes('svgBlip'))
  t.ok(slide.includes('<a:stretch><a:fillRect/></a:stretch>'))
  t.absent(out['ppt/media/image9.svg'])
  t.absent(decode(out['ppt/slides/_rels/slide3.xml.rels']).includes('image9.svg'))
})

test('setNotes writes existing notes and creates notes on a new slide', (t) => {
  const deck = open()
  deck.setNotes(1, 'Ask about sunsets first.')
  deck.setNotes(0, ['Open with the question', 'Pause'])
  const again = reopen(deck)
  const slides = again.read()
  t.is(slides[1].notes, 'Ask about sunsets first.')
  t.is(slides[0].notes, 'Open with the question\nPause')
  t.ok(files(again)['ppt/notesSlides/notesSlide2.xml'])
})

test('setNotes creates a notes master on a deck that has none', (t) => {
  const deck = pptx.create()
  deck.addSlide({ title: 'One' })
  deck.setNotes(0, 'Speaker notes')
  const again = reopen(deck)
  t.is(again.read()[0].notes, 'Speaker notes')
  const out = files(again)
  t.ok(out['ppt/notesMasters/notesMaster1.xml'])
  t.ok(out['ppt/theme/theme2.xml'])
  t.ok(decode(out['ppt/presentation.xml']).includes('<p:notesMasterIdLst><p:notesMasterId r:id='))
  const order = decode(out['ppt/presentation.xml'])
  t.ok(order.indexOf('notesMasterIdLst') < order.indexOf('sldIdLst'))
})

test('setText estimates overflow and can shrink text to fit', (t) => {
  const deck = open()
  const long = Array.from(
    { length: 14 },
    (_, i) => `Point ${i + 1} explains one more detail of the scattering story`
  )
  t.is(deck.setText({ slide: 1, shape: 'body' }, 'Short').overflow, false)
  const over = deck.setText({ slide: 1, shape: 'body' }, long)
  t.is(over.overflow, true)
  t.absent('fontScale' in over)
  const shrunk = deck.setText({ slide: 1, shape: 'body' }, long.slice(0, 13), { fit: 'shrink' })
  t.ok(shrunk.fontScale !== undefined && shrunk.fontScale < 1 && shrunk.fontScale >= 0.25)
  const xml = decode(files(deck)['ppt/slides/slide2.xml'])
  t.ok(/<a:bodyPr><a:normAutofit fontScale="\d+" lnSpcReduction="\d+"\/><\/a:bodyPr>/.test(xml))
  t.exception(
    () => deck.setText({ slide: 1, shape: 'body' }, 'x', bad({ fit: 'grow' })),
    /fit must be "shrink"/
  )
})
