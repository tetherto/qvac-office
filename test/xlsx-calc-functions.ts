import test from 'brittle'
import * as xlsx from '../lib/xlsx/index.ts'
import type { CellValue, Cells } from '../lib/xlsx/index.ts'

const DATA: Cells = {
  A1: 1,
  A2: 2,
  A3: '3',
  A4: true,
  A5: 'apple',
  A6: 'Banana',
  B1: 10,
  B2: 20,
  B3: 30,
  B4: 40,
  B5: 50,
  B6: 60,
  C1: '  some   spaced  text ',
  D1: 'x',
  D2: 'y',
  D4: 'z'
}

interface Alike {
  alike(actual: unknown, expected: unknown, message?: string): void
}

function evaluate(formula: string): CellValue {
  const book = xlsx.create()
  book.setCells(0, { ...DATA, Z100: formula })

  return book.get(0, 'Z100').value
}

function check(t: Alike, cases: [string, CellValue][]): void {
  for (const [formula, expected] of cases) {
    t.alike(evaluate(formula), expected, formula)
  }
}

test('arithmetic follows Excel precedence, coercion and errors', (t) => {
  check(t, [
    ['=1+2*3', 7],
    ['=(1+2)*3', 9],
    ['=-2^2', 4],
    ['=2^3^2', 64],
    ['=2^-1', 0.5],
    ['=50%', 0.5],
    ['=-10%*2', -0.2],
    ['=A1+A3', 4],
    ['=A1+A4', 2],
    ['=A1+E9', 1],
    ['=E9', 0],
    ['="3"*"4"', 12],
    ['=A5+1', '#VALUE!'],
    ['=1/0', '#DIV/0!'],
    ['=0^0', '#NUM!'],
    ['=(-8)^0.5', '#NUM!'],
    ['=#N/A+1', '#N/A'],
    ['=#REF!', '#REF!'],
    ['="a"&1&TRUE', 'a1TRUE'],
    ['=1/3&""', '0.333333333333333'],
    ['=10^15&""', '1E+15'],
    ['=0.1+0.2=0.3', true],
    ['=1<"a"', true],
    ['="a"<TRUE', true],
    ['="ABC"="abc"', true],
    ['=E9=0', true],
    ['=E9=""', true],
    ['=A1<>A2', true],
    ['=A1>=1', true]
  ])
})

test('text without digits before an exponent is not a number', (t) => {
  check(t, [
    ['=-"E5"', '#VALUE!'],
    ['="E5"%', '#VALUE!'],
    ['="E5"*1', '#VALUE!'],
    ['=-"1e400"', '#VALUE!'],
    ['=".5"*1', 0.5],
    ['="5."*1', 5],
    ['="-.5E1"*1', -5],
    ['="1,000%"*1', 10]
  ])

  const book = xlsx.create()
  book.setCells(0, { A1: 'E5', B1: '=-A1', C1: '=A1%' })
  t.is(book.get(0, 'B1').value, '#VALUE!')
  t.is(book.get(0, 'C1').value, '#VALUE!')
})

test('SUM and friends ignore text and booleans in ranges but coerce them as arguments', (t) => {
  check(t, [
    ['=SUM(A1:A6)', 3],
    ['=SUM(A1:A6,"3",TRUE)', 7],
    ['=SUM("x")', '#VALUE!'],
    ['=SUM(B:B)', 210],
    ['=SUM(1:1)', 11],
    ['=SUM(A1:B2,B3)', 63],
    ['=AVERAGE(A1:A6)', 1.5],
    ['=AVERAGE(E1:E5)', '#DIV/0!'],
    ['=MIN(B1:B6)', 10],
    ['=MAX(B1:B6,100)', 100],
    ['=MAX(E1:E3)', 0],
    ['=COUNT(A1:A6)', 2],
    ['=COUNT(A1:A6,"5",TRUE,"x")', 4],
    ['=COUNTA(A1:A6)', 6],
    ['=COUNTA(D1:D4)', 3],
    ['=COUNTBLANK(D1:D4)', 1],
    ['=COUNTBLANK(D1:D5)', 2],
    ['=SUM(A1,{1,2;3,4})', 11],
    ['=SUM(B1:B3*2)', null]
  ])
})

test('logical functions and lazy IF', (t) => {
  check(t, [
    ['=IF(A1=1,"one","other")', 'one'],
    ['=IF(A1=2,"one")', false],
    ['=IF(A1=1,,"x")', 0],
    ['=IF(A2>1,B2/A2,1/0)', 10],
    ['=IF(FALSE,1/0,"safe")', 'safe'],
    ['=IF("x",1,2)', '#VALUE!'],
    ['=IFERROR(1/0,"caught")', 'caught'],
    ['=IFERROR(5,"caught")', 5],
    ['=IFNA(#N/A,"na")', 'na'],
    ['=AND(TRUE,A1=1)', true],
    ['=AND(A1:A6)', true],
    ['=AND(E1:E3)', '#VALUE!'],
    ['=OR(FALSE,0)', false],
    ['=OR(A1>5,B1>5)', true],
    ['=NOT(A4)', false],
    ['=TRUE()', true]
  ])
})

test('math functions round the way Excel does', (t) => {
  check(t, [
    ['=ROUND(2.345,2)', 2.35],
    ['=ROUND(1.005,2)', 1.01],
    ['=ROUND(-2.5,0)', -3],
    ['=ROUND(1234.5,-2)', 1200],
    ['=ROUNDUP(1.201,2)', 1.21],
    ['=ROUNDUP(-1.201,1)', -1.3],
    ['=ROUNDDOWN(1.299,1)', 1.2],
    ['=ABS(-4)', 4],
    ['=INT(-1.5)', -2],
    ['=MOD(10,3)', 1],
    ['=MOD(-1,3)', 2],
    ['=MOD(1,0)', '#DIV/0!'],
    ['=POWER(2,10)', 1024],
    ['=SQRT(16)', 4],
    ['=SQRT(-1)', '#NUM!']
  ])
})

test('text functions', (t) => {
  check(t, [
    ['=CONCAT(A1:A3,"-",D1:D4)', '123-xyz'],
    ['=CONCATENATE("a",A1,"b")', 'a1b'],
    ['=_xlfn.CONCAT("a","b")', 'ab'],
    ['=TEXTJOIN(",",TRUE,D1:D4)', 'x,y,z'],
    ['=TEXTJOIN(",",FALSE,D1:D4)', 'x,y,,z'],
    ['=LEN(A5)', 5],
    ['=LEN(123.5)', 5],
    ['=LEFT(A5,3)', 'app'],
    ['=LEFT(A5)', 'a'],
    ['=RIGHT(A5,2)', 'le'],
    ['=MID(A5,2,3)', 'ppl'],
    ['=MID(A5,0,1)', '#VALUE!'],
    ['=UPPER(A5)', 'APPLE'],
    ['=LOWER(A6)', 'banana'],
    ['=TRIM(C1)', 'some spaced text']
  ])
})

test('criteria functions match numbers, text, wildcards and blanks', (t) => {
  check(t, [
    ['=COUNTIF(B1:B6,">25")', 4],
    ['=COUNTIF(B1:B6,30)', 1],
    ['=COUNTIF(A1:A6,3)', 1],
    ['=COUNTIF(A5:A6,"b*")', 1],
    ['=COUNTIF(A5:A6,"?pple")', 1],
    ['=COUNTIF(D1:D4,"")', 1],
    ['=COUNTIF(D1:D4,"<>")', 3],
    ['=COUNTIF(D1:D4,"<>x")', 3],
    ['=SUMIF(B1:B6,">=30")', 180],
    ['=SUMIF(A1:A6,"apple",B1:B6)', 50],
    ['=SUMIF(D1:D4,"",B1:B4)', 30],
    ['=AVERAGEIF(B1:B6,"<35")', 20],
    ['=AVERAGEIF(B1:B6,">99")', '#DIV/0!']
  ])
})

test('lookups find exact and approximate matches', (t) => {
  check(t, [
    ['=VLOOKUP(2,A1:B3,2,FALSE)', 20],
    ['=VLOOKUP("banana",A5:B6,2,FALSE)', 60],
    ['=VLOOKUP("b*",A5:B6,2,0)', 60],
    ['=VLOOKUP(9,A1:B3,2,FALSE)', '#N/A'],
    ['=VLOOKUP(2.5,A1:B2,2)', 20],
    ['=VLOOKUP(0,A1:B2,2,TRUE)', '#N/A'],
    ['=VLOOKUP(1,A1:B2,3,FALSE)', '#REF!'],
    ['=VLOOKUP(1,A1:B2,0,FALSE)', '#VALUE!'],
    ['=HLOOKUP(20,B2:B3,2,FALSE)', 30],
    ['=HLOOKUP(10,{10,20;"a","b"},2,FALSE)', 'a'],
    ['=MATCH(30,B1:B6,0)', 3],
    ['=MATCH(35,B1:B6)', 3],
    ['=MATCH(35,{60,50,40,30},-1)', 3],
    ['=MATCH("x",D1:D4,0)', 1],
    ['=INDEX(B1:B6,4)', 40],
    ['=INDEX(A1:B6,2,2)', 20],
    ['=INDEX(B1:B6,9)', '#REF!'],
    ['=SUM(INDEX(A1:B6,0,2))', 210],
    ['=INDEX(B1:B6,MATCH("apple",A1:A6,0))', 50]
  ])
})

test('volatile, unknown and spilling formulas are left uncomputed', (t) => {
  check(t, [
    ['=TODAY()', null],
    ['=NOW()+1', null],
    ['=TEXT(A1,"0")', null],
    ['=A1:A2', null],
    ['=Table1[Col]', null],
    ['=[1]Sheet1!A1', null],
    ['=MATCH(2,{3,1,2})', null],
    ['=UnknownName+1', null]
  ])
})
