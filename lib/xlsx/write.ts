import { Package } from '../opc/package.ts'
import * as xml from '../xml.ts'
import { cell } from './refs.ts'

export type EmbeddedValue = string | number | null | undefined

const SML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
const TYPES = {
  workbook: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  worksheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
}
const STYLES = `<styleSheet xmlns="${SML}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`

function escape(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function cellXml(value: EmbeddedValue, ref: string): string {
  if (value === null || value === undefined || value === '') return ''
  if (typeof value === 'number') return `<c r="${ref}"><v>${value}</v></c>`
  return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escape(value)}</t></is></c>`
}

function sheetXml(rows: EmbeddedValue[][]): string {
  const data = rows
    .map(
      (row, r) =>
        `<row r="${r + 1}">${row.map((value, c) => cellXml(value, cell(r, c))).join('')}</row>`
    )
    .join('')
  return `<worksheet xmlns="${SML}"><sheetData>${data}</sheetData></worksheet>`
}

export function writeWorkbook(rows: EmbeddedValue[][], sheet = 'Sheet1'): Uint8Array {
  const pkg = Package.create()
  const workbook = `<workbook xmlns="${SML}" xmlns:r="${xml.ns.r}"><sheets><sheet name="${escape(sheet)}" sheetId="1" r:id="rId1"/></sheets></workbook>`
  pkg.add('/xl/workbook.xml', xml.parse(workbook), TYPES.workbook)
  pkg.add('/xl/worksheets/sheet1.xml', xml.parse(sheetXml(rows)), TYPES.worksheet)
  pkg.add('/xl/styles.xml', xml.parse(STYLES), TYPES.styles)
  pkg.rels('/').add(REL_BASE + 'officeDocument', '/xl/workbook.xml')
  pkg.rels('/xl/workbook.xml').add(REL_BASE + 'worksheet', '/xl/worksheets/sheet1.xml')
  pkg.rels('/xl/workbook.xml').add(REL_BASE + 'styles', '/xl/styles.xml')
  return pkg.save()
}
