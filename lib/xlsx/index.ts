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
  InsertColumnsOptions,
  ColumnWidthOptions,
  AddSheetOptions,
  ImageOptions,
  SeriesRange,
  SheetChartSpec
} from './workbook.ts'
export type { CellStyle, CellAlign } from './styles.ts'
export type { ColumnSpan } from './columns.ts'
export type { Merged, Unmerged } from './merge.ts'
export type {
  RowSpanInput,
  ColumnSpanInput,
  RowHeightOptions,
  RowHeight,
  HideOptions,
  Hidden,
  FreezeOptions,
  Frozen
} from './layout.ts'
export type { AutoFilter } from './filter.ts'
export type { SortKey, SortOptions, Sorted } from './sort.ts'
export type {
  CellIsOperator,
  RuleValue,
  CellIsRule,
  ColorScaleRule,
  DataBarRule,
  TextRule,
  ConditionalRule,
  ConditionalAdded,
  ConditionalDeleted
} from './conditional.ts'
export type { DiffStyle } from './dxf.ts'
export type {
  ValidationType,
  ValidationOperator,
  ValidationMessage,
  ValidationError,
  ValidationRule,
  ValidationAdded,
  ValidationDeleted
} from './validation.ts'
export type { DefinedName, NameOptions, NameDefined, NameDeleted } from './names.ts'
export type { TableInfo, TableOptions, AddedTable, DeletedTable } from './tables.ts'
export type { FindOptions, Found } from './search.ts'
export type { HeaderGuess, SheetOutline } from './outline.ts'
export type { DocumentConversion } from '../opc/templates.ts'

export function open(bytes: Uint8Array | ArrayBuffer, opts?: OpenOptions): Workbook {
  return Workbook.open(bytes, opts)
}

export function create(opts?: CreateOptions): Workbook {
  return Workbook.create(opts)
}
