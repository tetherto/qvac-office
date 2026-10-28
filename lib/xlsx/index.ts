import { Workbook } from './workbook.ts'
import type { OpenOptions } from '../opc/package.ts'
import type { CreateOptions } from './template.ts'

export { Workbook }
export type {
  SheetRef,
  SheetInfo,
  SheetChart,
  WorkbookRead,
  Placed,
  RowSpan,
  Changed,
  Renamed,
  CellCount,
  Deleted,
  SheetCount,
  AddedSheet,
  CellValue,
  CellRead,
  CreateOptions,
  ReadOptions,
  CellOptions,
  CellInput,
  Cells,
  InsertRowsOptions,
  DeleteRowsOptions,
  DeleteColumnsOptions,
  ColumnWidthOptions,
  AddSheetOptions,
  ImageOptions,
  SeriesRange,
  SheetChartSpec
} from './workbook.ts'
export type { CellStyle, CellAlign } from './styles.ts'

export function open(bytes: Uint8Array | ArrayBuffer, opts?: OpenOptions): Workbook {
  return Workbook.open(bytes, opts)
}

export function create(opts?: CreateOptions): Workbook {
  return Workbook.create(opts)
}
