import test from 'brittle'
import * as xlsx from '../lib/xlsx/index.ts'
import * as utf8 from '../lib/utf8.ts'
import { fixture, decode, unzip, zipFiles } from './helpers/engine.ts'

const SHEET1 = 'xl/worksheets/sheet1.xml'
const SHEET2 = 'xl/worksheets/sheet2.xml'

function rewrite(bytes: Uint8Array, part: string, edit: (text: string) => string): Uint8Array {
  const files = unzip(bytes)
  files[part] = utf8.encode(edit(decode(files[part])))

  return zipFiles(files)
}

function cellXml(bytes: Uint8Array, part: string, ref: string): string {
  const found = new RegExp(`<c r="${ref}"[^>]*>.*?</c>`).exec(decode(unzip(bytes)[part]))

  return found ? found[0] : ''
}

test('setCells formulas read back computed before and after save', (t) => {
  const book = xlsx.create({ sheets: ['Data'] })
  book.setCells('Data', {
    A1: 2,
    A2: 3,
    A3: '=A1*A2',
    A4: '=A3&" items"',
    A5: '=A3>5',
    A6: '=A1/0'
  })
  t.is(book.get('Data', 'A3').value, 6)
  t.alike(book.read('Data').rows.flat(), [2, 3, 6, '6 items', true, '#DIV/0!'])
  book.setCells('Data', { A1: 5 })
  t.is(book.get('Data', 'A3').value, 15)

  const saved = book.save()
  t.is(cellXml(saved, SHEET1, 'A3'), '<c r="A3"><f>A1*A2</f><v>15</v></c>')
  t.is(cellXml(saved, SHEET1, 'A4'), '<c r="A4" t="str"><f>A3&amp;" items"</f><v>15 items</v></c>')
  t.is(cellXml(saved, SHEET1, 'A5'), '<c r="A5" t="b"><f>A3&gt;5</f><v>1</v></c>')
  t.is(cellXml(saved, SHEET1, 'A6'), '<c r="A6" t="e"><f>A1/0</f><v>#DIV/0!</v></c>')
  t.ok(decode(unzip(saved)['xl/workbook.xml']).includes('fullCalcOnLoad="1"'))
  t.is(xlsx.open(saved).get('Data', 'A4').value, '15 items')
})

test('formulas evaluate in dependency order across sheets', (t) => {
  const book = xlsx.create({ sheets: ['Inputs', 'Q1 Report'] })
  book.setCells('Q1 Report', { A1: '=A2+1', A2: '=A3*2', A3: '=Inputs!B2+SUM(Inputs!B1:B3)' })
  book.setCells('Inputs', {
    B1: 1,
    B2: 2,
    B3: 3,
    C1: "='Q1 Report'!A1*10",
    C2: "=SUM('q1 report'!A1:A3)"
  })
  const report = book.calculate()
  t.alike(report, { computed: 5, skipped: 0, cycles: [], updated: 5 })
  t.alike(book.read('Q1 Report').rows.flat(), [17, 16, 8])
  t.is(book.get('Inputs', 'C1').value, 170)
  t.is(book.get('Inputs', 'C2').value, 41)
})

test('cycles are reported and left uncomputed, and so are their dependents', (t) => {
  const book = xlsx.create({ sheets: ['S'] })
  book.setCells('S', {
    A1: '=B1+1',
    B1: '=A1+1',
    C1: '=A1*2',
    D1: '=SUM(D1:D2)',
    E1: 4,
    F1: '=E1*2'
  })
  const report = book.calculate()
  t.alike(report.cycles.sort(), ['S!A1', 'S!B1', 'S!D1'])
  t.is(report.skipped, 1)
  t.is(report.computed, 1)
  t.alike(book.read('S').rows[0], [null, null, null, null, 4, 8])
  t.is(cellXml(book.save(), SHEET1, 'A1'), '<c r="A1"><f>B1+1</f></c>')
})

test('sheets whose values did not change save byte for byte', (t) => {
  const first = xlsx.create({ sheets: ['One', 'Two'] })
  first.setCells('One', { A1: 1, A2: '=A1+1' })
  first.setCells('Two', { A1: 5, A2: '=A1*3', A3: '=SUM(A1:A2)' })
  const before = first.save()

  const book = xlsx.open(before)
  book.setCells('One', { A1: 10 })
  const after = book.save()
  t.alike(unzip(after)[SHEET2], unzip(before)[SHEET2])
  t.is(cellXml(after, SHEET1, 'A2'), '<c r="A2"><f>A1+1</f><v>11</v></c>')

  const untouched = xlsx.open(before)
  t.alike(untouched.calculate(), { computed: 3, skipped: 0, cycles: [], updated: 0 })
  const resaved = untouched.save()
  t.alike(unzip(resaved), unzip(before))
})

test('a formula the engine cannot compute keeps its cached value', (t) => {
  const base = xlsx.create({ sheets: ['S'] })
  base.setCells('S', { A1: 1, B1: '=TEXT(A1,"0.0")', C1: '=TODAY()', D1: '=B1&"!"' })
  const cachedXml = rewrite(base.save(), SHEET1, (text) =>
    text
      .replace('<f>TEXT(A1,"0.0")</f>', '<f>TEXT(A1,"0.0")</f><v>1.0</v>')
      .replace('<c r="B1">', '<c r="B1" t="str">')
  )
  const book = xlsx.open(cachedXml)
  book.setCells('S', { A1: 2 })
  t.alike(book.calculate(), { computed: 0, skipped: 3, cycles: [], updated: 0 })
  t.is(book.get('S', 'B1').value, '1.0')
  t.is(book.get('S', 'D1').value, null)
  t.ok(decode(unzip(book.save())['xl/workbook.xml']).includes('fullCalcOnLoad="1"'))
})

test('shared formulas are computed per cell and array formulas keep their cache', (t) => {
  const base = xlsx.create({ sheets: ['S'] })
  base.setCells('S', { A1: 1, A2: 2, A3: 3, B1: 0, B2: 0, B3: 0, C1: 0 })
  const shared = rewrite(base.save(), SHEET1, (text) =>
    text
      .replace(
        '<c r="B1"><v>0</v></c>',
        '<c r="B1"><f t="shared" ref="B1:B3" si="0">A1*10</f><v>0</v></c>'
      )
      .replace('<c r="B2"><v>0</v></c>', '<c r="B2"><f t="shared" si="0"/><v>0</v></c>')
      .replace('<c r="B3"><v>0</v></c>', '<c r="B3"><f t="shared" si="0"/><v>0</v></c>')
      .replace(
        '<c r="C1"><v>0</v></c>',
        '<c r="C1"><f t="array" ref="C1">SUM(A1:A3*2)</f><v>99</v></c>'
      )
  )
  const book = xlsx.open(shared)
  t.alike(book.calculate(), { computed: 3, skipped: 1, cycles: [], updated: 3 })
  t.alike(book.read('S', { range: 'B1:C3' }).rows, [
    [10, 99],
    [20, null],
    [30, null]
  ])
  t.is(cellXml(book.save(), SHEET1, 'B2'), '<c r="B2"><f t="shared" si="0"/><v>20</v></c>')
})

test('defined names resolve to their ranges and constants', (t) => {
  const base = xlsx.create({ sheets: ['S'] })
  base.setCells('S', { A1: 4, A2: 6, B1: '=SUM(Values)*Rate' })
  const named = rewrite(base.save(), 'xl/workbook.xml', (text) =>
    text.replace(
      '</sheets>',
      '</sheets><definedNames><definedName name="Values">S!$A$1:$A$2</definedName><definedName name="Rate">0.5</definedName></definedNames>'
    )
  )
  const book = xlsx.open(named)
  book.setCells('S', { A2: 16 })
  t.is(book.get('S', 'B1').value, 10)
})

test('calculate: false leaves formulas uncomputed until calculate() is called', (t) => {
  const book = xlsx.open(fixture('shared.xlsx'), { calculate: false })
  t.is(book.autoCalculate, false)
  book.setCells(0, { B2: 1000, C1: '=B4*2' })
  t.is(book.get(0, 'B4').value, 1650.5)
  t.is(book.get(0, 'C1').value, null)
  t.alike(book.calculate(), { computed: 2, skipped: 0, cycles: [], updated: 2 })
  t.is(book.get(0, 'C1').value, 2901)

  const later = xlsx.create()
  later.autoCalculate = false
  later.setCells(0, { A1: '=1+1' })
  t.is(xlsx.open(later.save()).get(0, 'A1').value, null)
})
