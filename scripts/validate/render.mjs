import fs from 'fs'
import path from 'path'
import process from 'process'
import pdfium from 'bare-pdfium'

const dir = process.argv[2] ?? 'rendered'

function inked(data) {
  for (let i = 0; i < data.length; i += 4) {
    if (data[i] < 250 || data[i + 1] < 250 || data[i + 2] < 250) return true
  }

  return false
}

let blankFiles = 0
for (const name of fs
  .readdirSync(dir)
  .filter((n) => n.endsWith('.pdf'))
  .sort()) {
  const doc = pdfium.open(new Uint8Array(fs.readFileSync(path.join(dir, name))))
  const count = doc.pageCount()
  const blank = []
  for (let page = 0; page < count; page++) {
    if (!inked(doc.render(page, { scale: 0.5 }).data)) {
      blank.push(page)
    }
  }

  doc.close()
  if (count === 0 || blank.length === count) {
    blankFiles++
  }

  console.log(`${name}: ${count} page(s)${blank.length ? `, blank: ${blank.join(', ')}` : ''}`)
}

if (blankFiles) {
  throw new Error(`${blankFiles} converted file(s) rendered blank`)
}
