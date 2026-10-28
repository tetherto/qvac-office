import * as xml from '../xml.ts'
import { DocsError } from '../errors.ts'
import type { Package } from '../opc/package.ts'

export interface CreateOptions {
  sheets?: string[]
  title?: string
  strict?: boolean
}

const { x: X, r: R } = xml.ns
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
const SML = 'application/vnd.openxmlformats-officedocument.spreadsheetml.'
const CORE_REL =
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'

function escape(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function sheetNames(input: unknown): string[] {
  const names = input === undefined ? ['Sheet1'] : input
  if (!Array.isArray(names) || names.length === 0) {
    throw new DocsError('INVALID_OPTIONS', 'sheets must be a non-empty array of sheet names')
  }

  const seen = new Set<string>()

  return names.map((name) => {
    if (
      typeof name !== 'string' ||
      name.length < 1 ||
      name.length > 31 ||
      /[:\\/?*[\]]/.test(name)
    ) {
      throw new DocsError('INVALID_NAME', `"${String(name)}" is not a valid sheet name`)
    }

    if (seen.has(name.toLowerCase())) {
      throw new DocsError('INVALID_NAME', `the sheet name "${name}" is repeated`)
    }

    seen.add(name.toLowerCase())

    return name
  })
}

function addSheets(pkg: Package, names: string[]): void {
  const sheets = names
    .map((name, i) => `<sheet name="${escape(name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
    .join('')
  pkg.add(
    '/xl/workbook.xml',
    xml.parse(
      `<workbook xmlns="${X}" xmlns:r="${R}"><workbookPr/><bookViews><workbookView activeTab="0"/></bookViews><sheets>${sheets}</sheets><calcPr calcId="191029"/></workbook>`
    ),
    SML + 'sheet.main+xml'
  )
  names.forEach((_, i) => {
    const part = `/xl/worksheets/sheet${i + 1}.xml`
    const selected = i === 0 ? ' tabSelected="1"' : ''
    pkg.add(
      part,
      xml.parse(
        `<worksheet xmlns="${X}" xmlns:r="${R}"><sheetViews><sheetView workbookViewId="0"${selected}/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`
      ),
      SML + 'worksheet+xml'
    )
    pkg.rels('/xl/workbook.xml').add(REL_BASE + 'worksheet', part)
  })
}

export function build(pkg: Package, opts: CreateOptions = {}): void {
  addSheets(pkg, sheetNames(opts.sheets))
  const title =
    typeof opts.title === 'string' && opts.title ? `<dc:title>${escape(opts.title)}</dc:title>` : ''
  pkg.add(
    '/docProps/core.xml',
    xml.parse(
      `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">${title}</cp:coreProperties>`
    ),
    'application/vnd.openxmlformats-package.core-properties+xml'
  )
  pkg.add(
    '/docProps/app.xml',
    xml.parse(
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>QVAC</Application></Properties>'
    ),
    'application/vnd.openxmlformats-officedocument.extended-properties+xml'
  )
  pkg.rels('/').add(REL_BASE + 'officeDocument', '/xl/workbook.xml')
  pkg.rels('/').add(CORE_REL, '/docProps/core.xml')
  pkg.rels('/').add(REL_BASE + 'extended-properties', '/docProps/app.xml')
}
