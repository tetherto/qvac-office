import * as xml from '../xml.ts'
import * as refs from './refs.ts'
import * as areas from './areas.ts'
import * as names from './names.ts'
import * as tables from './tables.ts'
import { SHEET_ORDER, type Sheet, type SheetEntry } from './sheet.ts'
import { DocsError } from '../errors.ts'
import type { Area } from './refs.ts'
import type { ReshapeHost } from './reshape.ts'

export interface AutoFilter {
  range: string | null
  changed: boolean
}

const X = xml.ns.x

function entryOf(host: ReshapeHost, sheet: Sheet): SheetEntry {
  const found = host.entries().find((e) => e.part === sheet.part)
  if (!found) {
    throw new DocsError('SHEET_NOT_FOUND', `sheet "${sheet.name}" is no longer in the workbook`)
  }

  return found
}

function guard(sheet: Sheet, area: Area): void {
  const ref = refs.range(area)
  const table = tables.on(sheet).find((t) => areas.intersects(t.area, area))
  if (table) {
    throw new DocsError(
      'TABLE_OVERLAP',
      `${ref} overlaps table ${table.name} at ${table.ref}, which has its own filter`
    )
  }
}

function clear(host: ReshapeHost, sheet: Sheet): AutoFilter {
  const existing = xml.child(sheet.root, X, 'autoFilter')
  if (!existing) {
    return { range: null, changed: false }
  }

  xml.remove(existing)
  names.setFilterDatabase(host, entryOf(host, sheet), null)
  sheet.touch()

  return { range: null, changed: true }
}

export function autofilter(host: ReshapeHost, sheet: Sheet, range: unknown): AutoFilter {
  if (range === null) {
    return clear(host, sheet)
  }

  const area = areas.parse(range)
  guard(sheet, area)
  const ref = refs.range(area)
  const existing = xml.child(sheet.root, X, 'autoFilter')
  if (existing?.getAttribute('ref') === ref) {
    return { range: ref, changed: false }
  }

  if (existing) {
    xml.remove(existing)
  }

  xml.insertOrdered(
    sheet.root,
    xml.setAttributes(xml.named(sheet.doc, 'autoFilter', X), { ref }),
    SHEET_ORDER
  )
  names.setFilterDatabase(host, entryOf(host, sheet), area)
  sheet.touch()

  return { range: ref, changed: true }
}
