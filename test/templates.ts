import test from 'brittle'
import * as docx from '../lib/docx/index.ts'
import * as pptx from '../lib/pptx/index.ts'
import * as xlsx from '../lib/xlsx/index.ts'
import * as utf8 from '../lib/utf8.ts'
import { decode, unzip, zipFiles } from './helpers/engine.ts'

const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'
const PML = 'application/vnd.openxmlformats-officedocument.presentationml.'
const SML = 'application/vnd.openxmlformats-officedocument.spreadsheetml.'

interface Retype {
  bytes: Uint8Array
  from: string
  to: string
}

function retype({ bytes, from, to }: Retype): Uint8Array {
  const files = unzip(bytes)
  const types = decode(files['[Content_Types].xml'])
  if (!types.includes(from)) {
    throw new Error(`no ${from} in the content types`)
  }

  files['[Content_Types].xml'] = utf8.encode(types.replace(`"${from}"`, `"${to}"`))

  return zipFiles(files)
}

function mainType(bytes: Uint8Array, part: string): string {
  const types = decode(unzip(bytes)['[Content_Types].xml'])
  const match = new RegExp(`PartName="${part}" ContentType="([^"]+)"`).exec(types)

  return match?.[1] ?? ''
}

test('a .dotx opens, reports template and saves as a .docx', (t) => {
  const bytes = retype({
    bytes: docx.create().save(),
    from: WML + 'document.main+xml',
    to: WML + 'template.main+xml'
  })
  const doc = docx.open(bytes)
  t.is(doc.template, true)
  doc.appendParagraphs(['From the template'])
  t.alike(doc.toDocument(), { changed: true, contentType: WML + 'document.main+xml' })
  t.is(doc.template, false)
  t.alike(doc.toDocument(), { changed: false, contentType: WML + 'document.main+xml' })
  const saved = doc.save()
  t.is(mainType(saved, '/word/document.xml'), WML + 'document.main+xml')
  t.is(docx.open(saved).template, false)
})

test('a macro-enabled Word template becomes a macro-enabled document', (t) => {
  const bytes = retype({
    bytes: docx.create().save(),
    from: WML + 'document.main+xml',
    to: 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml'
  })
  const doc = docx.open(bytes)
  t.ok(doc.template)
  t.is(doc.toDocument().contentType, 'application/vnd.ms-word.document.macroEnabled.main+xml')
})

test('a .potx opens and saves as a .pptx', (t) => {
  const bytes = retype({
    bytes: pptx.create().save(),
    from: PML + 'presentation.main+xml',
    to: PML + 'template.main+xml'
  })
  const deck = pptx.open(bytes)
  t.is(deck.template, true)
  t.alike(deck.toDocument(), { changed: true, contentType: PML + 'presentation.main+xml' })
  const saved = deck.save()
  t.is(mainType(saved, deck.main), PML + 'presentation.main+xml')
  t.is(pptx.open(saved).template, false)
})

test('an .xltx opens and saves as an .xlsx', (t) => {
  const bytes = retype({
    bytes: xlsx.create().save(),
    from: SML + 'sheet.main+xml',
    to: SML + 'template.main+xml'
  })
  const book = xlsx.open(bytes)
  t.is(book.template, true)
  book.setCells(0, { A1: 'From the template' })
  t.alike(book.toDocument(), { changed: true, contentType: SML + 'sheet.main+xml' })
  const saved = book.save()
  t.is(mainType(saved, book.main), SML + 'sheet.main+xml')
  t.is(xlsx.open(saved).get(0, 'A1').value, 'From the template')
})

test('regular files are not templates and toDocument leaves them alone', (t) => {
  const doc = docx.open(docx.create().save())
  t.is(doc.template, false)
  t.is(doc.toDocument().changed, false)
  t.absent(doc.changed)
})
