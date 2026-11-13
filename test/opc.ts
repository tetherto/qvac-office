import test from 'brittle'
import { Package } from '../lib/opc/package.ts'
import { DocsError } from '../lib/errors.ts'
import * as paths from '../lib/opc/paths.ts'
import * as zip from '../lib/opc/zip.ts'
import { encode } from '../lib/utf8.ts'
import { fixture, decode, unzip, zipFiles } from './helpers/engine.ts'

const LIMITS = { maxEntries: 10000, maxSize: 2 ** 29, maxRatio: 1000 }

function codeOf(fn: () => unknown): string | null {
  try {
    fn()
  } catch (err) {
    return err instanceof DocsError ? err.code : 'NOT_DOCS_ERROR'
  }

  return null
}

function utf16(text: string): Uint8Array {
  const out = new Uint8Array(2 + text.length * 2)
  out.set([0xff, 0xfe])
  for (let i = 0; i < text.length; i++) {
    out[2 + i * 2] = text.charCodeAt(i)
  }

  return out
}

test('paths resolve and relativize between parts', (t) => {
  t.is(paths.resolve('/ppt/slides/slide1.xml', '../media/image1.png'), '/ppt/media/image1.png')
  t.is(paths.resolve('/', 'ppt/presentation.xml'), '/ppt/presentation.xml')
  t.is(paths.resolve('/ppt/presentation.xml', '/docProps/app.xml'), '/docProps/app.xml')
  t.is(
    paths.relative('/ppt/slides/slide4.xml', '/ppt/slideLayouts/slideLayout2.xml'),
    '../slideLayouts/slideLayout2.xml'
  )
  t.is(paths.relative('/ppt/presentation.xml', '/ppt/slides/slide1.xml'), 'slides/slide1.xml')
  t.is(paths.relative('/', '/ppt/presentation.xml'), 'ppt/presentation.xml')
  t.is(paths.relsFor('/ppt/slides/slide1.xml'), '/ppt/slides/_rels/slide1.xml.rels')
  t.is(paths.relsFor('/'), '/_rels/.rels')
})

test('an untouched package saves every part byte for byte', (t) => {
  const input = fixture('basic.pptx')
  const pkg = Package.open(input)
  pkg.xml('/ppt/presentation.xml')
  t.is(pkg.changed, false)
  const before = unzip(input)
  const after = unzip(pkg.save())
  t.alike(Object.keys(after).sort(), Object.keys(before).sort())
  for (const name of Object.keys(before)) {
    t.alike(after[name], before[name], name)
  }
})

test('content types resolve overrides before defaults', (t) => {
  const pkg = Package.open(fixture('basic.pptx'))
  t.is(
    pkg.contentType('/ppt/slides/slide1.xml'),
    'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
  )
  t.is(pkg.contentType('/ppt/media/image1.png'), 'image/png')
  t.is(pkg.contentType('/ppt/media/IMAGE9.PNG'), 'image/png')
})

test('relationships resolve targets and allocate free ids', (t) => {
  const pkg = Package.open(fixture('basic.pptx'))
  const rels = pkg.rels('/ppt/presentation.xml')
  t.is(rels.get('rId7')?.part, '/ppt/slides/slide1.xml')
  const id = rels.add('urn:test', '/ppt/slides/slide1.xml')
  t.is(id, 'rId11')
  t.ok(pkg.changed)
  t.is(Package.open(pkg.save()).rels('/ppt/presentation.xml').get(id)?.target, 'slides/slide1.xml')
})

test('prune drops parts no relationship reaches', (t) => {
  const pkg = Package.open(fixture('basic.pptx'))
  pkg.add('/ppt/orphan.xml', encode('<orphan/>'), 'application/xml')
  t.alike(pkg.prune(), ['/ppt/orphan.xml'])
  t.absent(pkg.has('/ppt/orphan.xml'))
  t.ok(pkg.has('/ppt/slides/slide1.xml'))
})

test('part names match whatever their case', (t) => {
  const files = unzip(fixture('basic.pptx'))
  const relsName = 'ppt/slides/_rels/slide3.xml.rels'
  files[relsName] = encode(decode(files[relsName]).replace('media/image1.png', 'media/IMAGE1.png'))
  const pkg = Package.open(zipFiles(files))
  const image = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
  t.alike(pkg.related('/ppt/slides/slide3.xml', image), ['/ppt/media/image1.png'])
  t.ok(pkg.has('/ppt/media/IMAGE1.png'))
  t.alike(pkg.bytes('/PPT/MEDIA/image1.png'), files['ppt/media/image1.png'])
  t.is(pkg.freeName('/ppt/media/Image%d.png'), '/ppt/media/Image2.png')
  t.is(
    codeOf(() => pkg.add('/ppt/media/IMAGE1.PNG', new Uint8Array(1))),
    'PART_EXISTS'
  )
  pkg.add('/ppt/orphan.xml', encode('<orphan/>'), 'application/xml')
  t.alike(pkg.prune(), ['/ppt/orphan.xml'])
  t.ok(unzip(pkg.save())['ppt/media/image1.png'])
})

test('removing a part drops its relationships and override', (t) => {
  const pkg = Package.open(fixture('basic.pptx'))
  pkg.remove('/ppt/notesSlides/notesSlide1.xml')
  const saved = unzip(pkg.save())
  t.absent(saved['ppt/notesSlides/notesSlide1.xml'])
  t.absent(saved['ppt/notesSlides/_rels/notesSlide1.xml.rels'])
  t.absent(decode(saved['[Content_Types].xml']).includes('notesSlide1.xml'))
})

test('rejects input that is not an Office package', (t) => {
  t.exception(() => Package.open(encode('hello')), /not a readable zip/)
  t.exception(() => Package.open(zipFiles({ 'a.txt': encode('x') })), /no \[Content_Types\]\.xml/)
  const text: unknown = 'not bytes'
  t.exception(
    () => Package.open(text instanceof Uint8Array ? text : encode(String(text)).subarray(0, 0)),
    /not a readable zip/
  )
})

test('refuses XML with a DOCTYPE', (t) => {
  const bytes = zipFiles({
    '[Content_Types].xml': encode('<!DOCTYPE x [<!ENTITY a "b">]><Types/>')
  })
  t.is(
    codeOf(() => Package.open(bytes)),
    'UNSAFE_XML'
  )
})

test('enforces entry and size limits', (t) => {
  const bytes = zipFiles({
    '[Content_Types].xml': encode('<Types/>'),
    'a.bin': new Uint8Array(4096)
  })
  t.exception(() => Package.open(bytes, { limits: { maxEntries: 1 } }), /more than 1 entries/)
  t.exception(() => Package.open(bytes, { limits: { maxSize: 1024 } }), /expands past 1024 bytes/)
})

test('refuses XML nested deeper than the limit', (t) => {
  const deep = '<a>'.repeat(300) + '</a>'.repeat(300)
  const bytes = zipFiles({ '[Content_Types].xml': encode(`<Types>${deep}</Types>`) })
  t.exception(() => Package.open(bytes), /more than 256 levels/)
  t.is(
    codeOf(() => Package.open(bytes)),
    'UNSAFE_XML'
  )
})

test('nesting hidden behind comments, CDATA or attribute values still counts', (t) => {
  const nested = (open: string): Uint8Array =>
    zipFiles({
      '[Content_Types].xml': encode(`<Types>${open.repeat(300)}${'</a>'.repeat(300)}</Types>`)
    })
  t.is(
    codeOf(() => Package.open(nested('<a><!--</a>-->'))),
    'UNSAFE_XML'
  )
  t.is(
    codeOf(() => Package.open(nested('<a><![CDATA[</a>]]>'))),
    'UNSAFE_XML'
  )
  t.is(
    codeOf(() => Package.open(nested('<a><?pi </a>?>'))),
    'UNSAFE_XML'
  )
  t.is(
    codeOf(() => Package.open(nested('<a x="/>">'))),
    'UNSAFE_XML'
  )
  t.is(
    codeOf(() => Package.open(nested("<a x='</a>'>"))),
    'UNSAFE_XML'
  )
  const flat = '<a x="/>"><!--<b>--><![CDATA[<c>]]></a>'.repeat(300)
  t.execution(() =>
    Package.open(zipFiles({ '[Content_Types].xml': encode(`<Types>${flat}</Types>`) }))
  )
})

test('refuses an entry that expands far beyond its compressed size', (t) => {
  const bytes = zipFiles({
    '[Content_Types].xml': encode('<Types/>'),
    'bomb.bin': new Uint8Array(8 * 1024 * 1024)
  })
  t.exception(() => Package.open(bytes), /bomb\.bin expands more than 1000 times/)
  t.execution(() => Package.open(bytes, { limits: { maxRatio: 100000 } }))
})

test('stops inflating an entry once it passes its declared size', (t) => {
  const bytes = zip.write([
    { name: 'bomb.bin', data: new Uint8Array(8 * 1024 * 1024), compress: true }
  ])
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const central = view.getUint32(bytes.length - 6, true)
  view.setUint32(22, 100, true)
  view.setUint32(central + 24, 100, true)
  t.exception(() => zip.read(bytes, LIMITS), /bomb\.bin inflates past its declared size/)
  t.is(
    codeOf(() => zip.read(bytes, LIMITS)),
    'NOT_OOXML'
  )
})

test('refuses an entry that inflates short of its declared size', (t) => {
  const bytes = zip.write([{ name: 'a.txt', data: encode('hello hello hello'), compress: true }])
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const central = view.getUint32(bytes.length - 6, true)
  view.setUint32(central + 24, 100, true)
  t.exception(() => zip.read(bytes, LIMITS), /a\.txt has the wrong size/)
})

test('reads XML parts with a byte order mark or in UTF-16', (t) => {
  const types =
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>'
  const bom = new Uint8Array([0xef, 0xbb, 0xbf, ...encode(types)])
  t.is(
    Package.open(zipFiles({ '[Content_Types].xml': bom })).contentType('/a.xml'),
    'application/xml'
  )
  t.is(
    Package.open(zipFiles({ '[Content_Types].xml': utf16(types) })).contentType('/a.xml'),
    'application/xml'
  )
})

test('a UTF-16 part is written back as UTF-8 with a matching declaration', (t) => {
  const types =
    '<?xml version="1.0" encoding="UTF-16"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
  const pkg = Package.open(zipFiles({ '[Content_Types].xml': utf16(types) }))
  pkg.contentTypes.addDefault('png', 'image/png')
  const out = decode(unzip(pkg.save())['[Content_Types].xml'])
  t.ok(out.startsWith('<?xml version="1.0" encoding="UTF-8"?>'))
  t.ok(out.includes('Extension="png"'))
})

test('the zip reader checks sizes, checksums, methods and encryption', (t) => {
  const good = zip.write([{ name: 'a.txt', data: encode('hello hello hello'), compress: true }])
  t.alike(
    zip.read(good, LIMITS).map((e) => decode(e.data)),
    ['hello hello hello']
  )
  const stored = zip.write([{ name: 'b.txt', data: encode('plain'), compress: false }])
  const flip = (bytes: Uint8Array, at: number, value: number): Uint8Array => {
    const copy = bytes.slice()
    copy[at] = value
    return copy
  }

  t.exception(() => zip.read(flip(stored, 56, stored[56] ^ 0xff), LIMITS), /fails its checksum/)
  t.exception(() => zip.read(flip(stored, 50, 12), LIMITS), /compression method 12/)
  t.exception(() => zip.read(flip(stored, 48, 1), LIMITS), /encrypted/)
  t.exception(() => zip.read(stored.subarray(0, 40), LIMITS), /not a readable zip archive/)
  t.alike(zip.read(zip.write([]), LIMITS), [])
})

test('zip crc32 matches the reference value', (t) => {
  t.is(zip.crc32(encode('The quick brown fox jumps over the lazy dog')), 0x414fa339)
})
