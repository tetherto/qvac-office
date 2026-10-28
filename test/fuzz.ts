import test from 'brittle'
import process from 'process'
import { pptx, docx, xlsx, DocsError } from '../lib/index.ts'
import { encode } from '../lib/utf8.ts'
import { errorMessage } from '../lib/guards.ts'
import { fixture, decode, unzip, zipFiles } from './helpers/engine.ts'

type Random = () => number

const ROUNDS = Number(process.env.FUZZ_ROUNDS ?? 60)
const JUNK = ['<x>', '</a:p>', '"', '&', '<!--', 'r:id="rId999"', '\u0000']

function random(seed: number): Random {
  let a = seed >>> 0
  return () => {
    a = (a + 0x6d2b79f5) >>> 0
    let t = a
    t = Math.imul(t ^ (t >>> 15), t | 1)
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61)
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

const targets: [string, (bytes: Uint8Array) => void][] = [
  [
    'basic.pptx',
    (b) => {
      const deck = pptx.open(b)
      deck.read()
      deck.setText({ slide: 0, shape: 'title' }, 'Fuzzed')
      deck.addSlide({ title: 'More' })
      deck.save()
    }
  ],
  [
    'charts.pptx',
    (b) => {
      const deck = pptx.open(b)
      deck.read()
      deck.setChartData(
        { slide: 0, shape: 3 },
        { categories: ['a'], series: [{ name: 's', values: [1] }] }
      )
      deck.save()
    }
  ],
  [
    'report.docx',
    (b) => {
      const doc = docx.open(b)
      doc.read()
      doc.appendParagraphs('Fuzzed')
      doc.setText(0, 'Title')
      doc.save()
    }
  ],
  [
    'sales.xlsx',
    (b) => {
      const book = xlsx.open(b)
      book.sheets()
      book.read(0)
      book.setCells(0, { A20: 'Fuzzed', B20: '=1+1' })
      book.save()
    }
  ],
  [
    'shared.xlsx',
    (b) => {
      const book = xlsx.open(b)
      book.read(0)
      book.appendRows(0, [['x', 1]])
      book.save()
    }
  ]
]

function pick<T>(items: T[], rand: Random): T | undefined {
  return items[Math.floor(rand() * items.length)]
}

function cut(text: string, match: RegExpMatchArray | undefined): string {
  if (!match || match.index === undefined) return text
  return text.slice(0, match.index) + text.slice(match.index + match[0].length)
}

function damageText(text: string, kind: number, rand: Random): string {
  const at = Math.floor(rand() * text.length)
  if (kind === 4) {
    return cut(text, pick([...text.matchAll(/ [A-Za-z:]+="[^"]*"/g)], rand))
  }

  if (kind === 5) {
    return cut(text, pick([...text.matchAll(/<[A-Za-z:]+(?: [^<>]*)?\/>/g)], rand))
  }

  if (kind === 2) {
    return text.slice(0, at) + text.slice(at + Math.floor(rand() * 200))
  }

  return text.slice(0, at) + (pick(JUNK, rand) ?? '') + text.slice(at)
}

function mutate(bytes: Uint8Array, rand: Random): Uint8Array {
  const kind = Math.floor(rand() * 6)
  if (kind === 0) {
    const out = bytes.slice()
    for (let i = 0; i < 8; i++) {
      out[Math.floor(rand() * out.length)] = Math.floor(rand() * 256)
    }

    return out
  }

  if (kind === 1) {
    return bytes.slice(0, Math.floor(rand() * bytes.length))
  }

  const files = unzip(bytes)
  const name = pick(
    Object.keys(files).filter((n) => n.endsWith('.xml') || n.endsWith('.rels')),
    rand
  )
  if (name) {
    files[name] = encode(damageText(decode(files[name]), kind, rand))
  }

  return zipFiles(files)
}

for (const [name, attempt] of targets) {
  test(`damaged ${name} fails only with DocsError`, (t) => {
    const source = fixture(name)
    const rand = random(name.length * 7919)
    const foreign: string[] = []
    for (let i = 0; i < ROUNDS; i++) {
      try {
        attempt(mutate(source, rand))
      } catch (err) {
        if (!(err instanceof DocsError)) {
          foreign.push(`round ${i}: ${errorMessage(err)}`)
        }
      }
    }

    t.alike(foreign, [])
  })
}
