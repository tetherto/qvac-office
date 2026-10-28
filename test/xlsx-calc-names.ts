import test from 'brittle'
import * as xlsx from '../lib/xlsx/index.ts'

test('a formula over a defined name follows the name when it changes', (t) => {
  const book = xlsx.create()
  book.setCells(0, { A1: 2, A2: 5, B1: '=Rate*10' })
  book.defineName('Rate', 'Sheet1!$A$1')
  t.is(book.get(0, 'B1').value, 20)
  book.defineName('Rate', 'Sheet1!$A$2')
  t.is(book.get(0, 'B1').value, 50)
  const reopened = xlsx.open(book.save())
  t.is(reopened.get(0, 'B1').value, 50)
})
