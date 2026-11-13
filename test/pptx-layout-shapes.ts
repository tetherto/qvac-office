import test from 'brittle'
import * as pptx from '../lib/pptx/index.ts'
import { DocsError } from '../lib/errors.ts'
import * as xml from '../lib/xml.ts'
import { fixture, decode, unzip } from './helpers/engine.ts'

const open = (): pptx.Presentation => pptx.open(fixture('basic.pptx'))
const reopen = (deck: pptx.Presentation): pptx.Presentation => pptx.open(deck.save())
const slideXml = (deck: pptx.Presentation, n: number): string =>
  decode(unzip(deck.save())[`ppt/slides/slide${n}.xml`])

function codeOf(fn: () => unknown): string | null {
  try {
    fn()
  } catch (err) {
    return err instanceof DocsError ? err.code : 'NOT_DOCS_ERROR'
  }

  return null
}

test('setFill and setLine write solid fills and outlines in schema order', (t) => {
  const deck = open()
  t.alike(deck.setFill({ slide: 0, shape: 'title' }, '#1a73e8'), { changed: true })
  t.alike(
    deck.setLine({ slide: 0, shape: 'title' }, { color: 'C0392B', width: '2pt', dash: 'dash' }),
    {
      changed: true
    }
  )
  const xml = slideXml(deck, 1)
  t.ok(xml.includes('<a:solidFill><a:srgbClr val="1A73E8"/></a:solidFill><a:ln w="25400">'))
  t.ok(
    xml.includes(
      '<a:ln w="25400"><a:solidFill><a:srgbClr val="C0392B"/></a:solidFill><a:prstDash val="dash"/></a:ln>'
    )
  )

  deck.setFill({ slide: 0, shape: 'title' }, null)
  deck.setLine({ slide: 0, shape: 'title' }, null)
  const cleared = slideXml(deck, 1)
  t.ok(cleared.includes('<a:noFill/><a:ln><a:noFill/></a:ln>'))
  t.alike(reopen(deck).read()[0].shapes[0].paragraphs, deck.read()[0].shapes[0].paragraphs)
})

test('setFill and setLine refuse what they cannot style', (t) => {
  const deck = open()
  t.is(
    codeOf(() => deck.setFill({ slide: 2, shape: 4 }, 'FF0000')),
    'NOT_SHAPE'
  )
  t.is(
    codeOf(() => deck.setFill({ slide: 0, shape: 2 }, 'red')),
    'INVALID_STYLE'
  )
  t.is(
    codeOf(() => deck.setLine({ slide: 0, shape: 2 }, { dash: 'wavy' as 'dash' })),
    'INVALID_STYLE'
  )
  t.is(
    codeOf(() => deck.setLine({ slide: 0, shape: 2 }, { width: '-1pt' })),
    'INVALID_STYLE'
  )
})

test('duplicateShape copies a picture to the same slide with an offset and a fresh id', (t) => {
  const deck = open()
  const copy = deck.duplicateShape({ slide: 2, shape: 3 })
  t.is(copy.id, 7)
  t.is(copy.slide, 2)
  t.alike([copy.x, copy.y], [914400 + 228600, 1828800 + 228600])
  t.ok(copy.overlaps.includes(3))
  const shapes = reopen(deck).read()[2].shapes
  t.alike(
    shapes.map((s) => s.kind),
    ['text', 'picture', 'table', 'group', 'text', 'picture']
  )
  t.ok(unzip(deck.save())['ppt/media/image1.png'])
})

test('duplicateShape copies a picture and a group to another slide', (t) => {
  const deck = open()
  const picture = deck.duplicateShape({ slide: 2, shape: 3 }, { slide: 0, x: '1in', y: 0 })
  t.alike([picture.slide, picture.x, picture.y], [0, 914400, 0])
  const group = deck.duplicateShape({ slide: 2, shape: 5 }, { slide: 1 })
  const again = reopen(deck).read()
  t.alike(
    again[0].shapes.map((s) => s.kind),
    ['text', 'text', 'picture']
  )
  t.alike(
    again[1].shapes.map((s) => [s.kind, s.group ?? null]),
    [
      ['text', null],
      ['text', null],
      ['group', null],
      ['text', group.id]
    ]
  )
  t.alike(again[1].shapes[3].paragraphs?.[0].text, 'Grouped caption')
  const rels = decode(unzip(deck.save())['ppt/slides/_rels/slide1.xml.rels'])
  t.ok(rels.includes('../media/image1.png'))
})

function turnedGroup(attributes: Record<string, string>): pptx.Presentation {
  const deck = open()
  deck.group(2, [3, 4])
  deck.setBox({ slide: 2, shape: 7 }, { y: 2743200 })
  const doc = deck.package.xml('/ppt/slides/slide3.xml')
  const group = xml.descendants(doc, xml.ns.p, 'grpSp')[0]
  const xfrm = xml.descendants(group, xml.ns.a, 'xfrm')[0]
  for (const [name, value] of Object.entries(attributes)) {
    xfrm.setAttribute(name, value)
  }

  return deck
}

function copyXfrm(deck: pptx.Presentation): string {
  return (
    slideXml(deck, 1)
      .match(/<a:xfrm[^>]*>/g)
      ?.at(-1) ?? ''
  )
}

test('duplicateShape places a child of a rotated group where it shows on the slide', (t) => {
  const deck = turnedGroup({ rot: '5400000' })
  const copy = deck.duplicateShape({ slide: 2, shape: 3 }, { slide: 0 })
  t.alike([copy.x, copy.y, copy.width, copy.height], [3200400, 457200, 1828800, 914400])
  t.ok(copyXfrm(deck).includes('rot="5400000"'))
})

test('duplicateShape places a child of a flipped group where it shows on the slide', (t) => {
  const deck = turnedGroup({ flipH: '1' })
  const copy = deck.duplicateShape({ slide: 2, shape: 3 }, { slide: 0 })
  t.alike([copy.x, copy.y], [5486400, 2743200])
  t.ok(copyXfrm(deck).includes('flipH="1"'))
})

test('overlaps sees a child of a rotated group where it shows on the slide', (t) => {
  const deck = turnedGroup({ rot: '5400000' })
  const box = deck.addTextBox(2, { x: 3200400, y: 457200, width: 100, height: 100 })
  t.alike(deck.setBox({ slide: 2, shape: 3 }, { x: 914400 }).overlaps, [2, box.id])
})

test('duplicateShape copies a chart with its own chart part and workbook', (t) => {
  const deck = pptx.create()
  deck.addSlide({ layout: 'Blank' })
  deck.addSlide({ layout: 'Blank' })
  const placed = deck.addChart(0, {
    type: 'bar',
    categories: ['A', 'B'],
    series: [{ name: 'S', values: [1, 2] }]
  })
  const copy = deck.duplicateShape({ slide: 0, shape: placed.id }, { slide: 1 })
  deck.setChartData(
    { slide: 1, shape: copy.id },
    { categories: ['X'], series: [{ name: 'T', values: [9] }] }
  )
  const again = reopen(deck).read()
  t.alike(again[0].shapes[0].chart?.categories, ['A', 'B'], 'the original keeps its data')
  t.alike(again[1].shapes[0].chart?.categories, ['X'])
  const files = unzip(deck.save())
  t.ok(files['ppt/charts/chart2.xml'])
  t.ok(files['ppt/embeddings/Microsoft_Excel_Worksheet2.xlsx'])
  t.ok(
    decode(files['ppt/charts/_rels/chart2.xml.rels']).includes('Microsoft_Excel_Worksheet2.xlsx')
  )
  t.ok(decode(files['[Content_Types].xml']).includes('/ppt/charts/chart2.xml'))
})

test('reorder moves a shape among its siblings', (t) => {
  const deck = open()
  t.alike(deck.reorder({ slide: 2, shape: 3 }, 'front'), { moved: true, position: 3 })
  t.alike(deck.reorder({ slide: 2, shape: 3 }, 'front'), { moved: false, position: 3 })
  t.alike(deck.reorder({ slide: 2, shape: 3 }, 'backward'), { moved: true, position: 2 })
  t.alike(deck.reorder({ slide: 2, shape: 4 }, 'back'), { moved: true, position: 0 })
  t.alike(deck.reorder({ slide: 2, shape: 2 }, 'forward'), { moved: true, position: 2 })
  t.alike(
    reopen(deck)
      .read()[2]
      .shapes.map((s) => s.id),
    [4, 3, 2, 5, 6]
  )
  t.is(
    codeOf(() => deck.reorder({ slide: 2, shape: 3 }, 'top' as 'front')),
    'INVALID_OPTIONS'
  )
})

test('group wraps shapes in a group with an identity transform, and ungroup releases them', (t) => {
  const deck = open()
  const grouped = deck.group(2, [4, 3], { name: 'Figures' })
  t.alike(grouped, {
    id: 7,
    x: 914400,
    y: 1828800,
    width: 3657600 + 3657600 - 914400,
    height: 914400,
    children: [3, 4]
  })
  const again = reopen(deck)
  const shapes = again.read()[2].shapes
  t.alike(
    shapes.map((s) => [s.id, s.group ?? null]),
    [
      [2, null],
      [7, null],
      [3, 7],
      [4, 7],
      [5, null],
      [6, 5]
    ]
  )
  t.ok(slideXml(deck, 3).includes('name="Figures"'))
  t.alike(again.ungroup({ slide: 2, shape: 7 }), { ungrouped: 7, children: [3, 4] })
  t.alike(
    reopen(again)
      .read()[2]
      .shapes.map((s) => [s.id, s.box]),
    [
      [2, { x: 457200, y: 274638, width: 8229600, height: 1143000 }],
      [3, { x: 914400, y: 1828800, width: 1828800, height: 914400 }],
      [4, { x: 3657600, y: 1828800, width: 3657600, height: 914400 }],
      [5, { x: 914400, y: 4572000, width: 2743200, height: 914400 }],
      [6, { x: 914400, y: 4572000, width: 2743200, height: 914400 }]
    ]
  )
})

test('ungroup maps children through a scaled group', (t) => {
  const deck = open()
  deck.group(2, [3, 4])
  deck.setBox({ slide: 2, shape: 7 }, { x: 0, y: 0, width: 6400800 * 2, height: 914400 * 2 })
  deck.ungroup({ slide: 2, shape: 7 })
  const boxes = reopen(deck)
    .read()[2]
    .shapes.filter((s) => s.id === 3 || s.id === 4)
    .map((s) => s.box)
  t.alike(boxes, [
    { x: 0, y: 0, width: 1828800 * 2, height: 914400 * 2 },
    { x: (3657600 - 914400) * 2, y: 0, width: 3657600 * 2, height: 914400 * 2 }
  ])
})

test('group refuses placeholders, shapes in different groups and lone shapes', (t) => {
  const deck = open()
  t.is(
    codeOf(() => deck.group(2, [2, 3])),
    'CANNOT_GROUP'
  )
  t.is(
    codeOf(() => deck.group(2, [3, 6])),
    'CANNOT_GROUP'
  )
  t.is(
    codeOf(() => deck.group(2, [3])),
    'CANNOT_GROUP'
  )
  t.is(
    codeOf(() => deck.group(2, [3, 3])),
    'CANNOT_GROUP'
  )
  t.is(
    codeOf(() => deck.ungroup({ slide: 2, shape: 3 })),
    'NOT_GROUP'
  )
})

test('the new shape edits take a slide by its id', (t) => {
  const deck = open()
  const id = deck.read()[2].id
  const copy = deck.duplicateShape({ slide: { id }, shape: 3 }, { slide: { id } })
  t.is(copy.slide, 2)
  t.is(deck.group({ id }, [3, copy.id]).children.length, 2)
  t.is(deck.setLayout({ id }, 'Title and Content').layout, 'Title and Content')
})
