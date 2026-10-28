import * as s from './schemas.ts'
import { group, opt, req, type OpTable, type Schema } from './spec.ts'

const CELL_STYLE: Schema = {
  type: 'object',
  properties: {
    format: { type: 'string', description: 'a number format such as "0.00%" or "yyyy-mm-dd"' },
    bold: s.BOOLEAN,
    italic: s.BOOLEAN,
    underline: s.BOOLEAN,
    color: s.COLOR,
    fill: s.COLOR,
    size: s.NUMBER,
    font: s.STRING,
    align: s.ALIGN,
    wrap: s.BOOLEAN
  },
  additionalProperties: false
}
const CELLS: Schema = {
  type: 'object',
  additionalProperties: s.CELL_VALUE,
  description: 'cells by A1 reference'
}
const ROWS: Schema = { type: 'array', items: { type: 'array', items: s.CELL_VALUE } }
const SHEET = req('sheet', s.SHEET)

export const XLSX: OpTable = {
  sheets: { summary: 'Sheets with their kind, state and used range', slots: [], query: true },
  read: {
    summary: 'Values, formulas, merges and charts of a range',
    slots: [SHEET, group(opt('range', s.RANGE), opt('limit', s.POSITIVE))],
    query: true
  },
  get: {
    summary: "One cell's value, formula and number format",
    slots: [SHEET, req('ref', s.CELL)],
    query: true
  },
  setCells: {
    summary: 'Write values, formulas, dates and styles by A1 reference',
    slots: [SHEET, req('cells', CELLS)]
  },
  format: {
    summary: 'Style a range',
    slots: [SHEET, req('range', s.RANGE), req('style', CELL_STYLE)]
  },
  appendRows: { summary: 'Append rows below the used range', slots: [SHEET, req('rows', ROWS)] },
  insertRows: {
    summary: 'Insert rows, moving every reference below',
    slots: [SHEET, group(req('at', s.POSITIVE), opt('rows', ROWS), opt('count', s.POSITIVE))]
  },
  deleteRows: {
    summary: 'Delete rows, moving every reference below',
    slots: [SHEET, group(req('at', s.POSITIVE), opt('count', s.POSITIVE))]
  },
  deleteColumns: {
    summary: 'Delete columns, moving every reference to the right',
    slots: [SHEET, group(req('column', s.COLUMN), opt('count', s.POSITIVE))]
  },
  setColumnWidth: {
    summary: "Set a column's width in characters",
    slots: [SHEET, group(req('column', s.COLUMN), req('width', s.NUMBER))]
  },
  addSheet: { summary: 'Add a sheet', slots: [req('name', s.STRING), group(opt('at', s.INDEX))] },
  renameSheet: {
    summary: 'Rename a sheet and the formulas that name it',
    slots: [SHEET, req('name', s.STRING)]
  },
  deleteSheet: { summary: 'Delete a sheet', slots: [SHEET] },
  addImage: {
    summary: 'Place an image at a cell',
    slots: [
      SHEET,
      { key: 'image', schema: s.IMAGE, required: true, image: true },
      group(
        opt('cell', s.CELL),
        opt('width', s.LENGTH),
        opt('height', s.LENGTH),
        opt('name', s.STRING),
        opt('description', s.STRING)
      )
    ]
  },
  addChart: {
    summary: 'Chart cell ranges, linked to the sheet',
    slots: [
      SHEET,
      group(
        req('type', s.CHART_TYPE),
        req('categories', s.RANGE),
        req('series', {
          type: 'array',
          items: {
            type: 'object',
            properties: { values: s.RANGE, name: s.STRING, nameCell: s.CELL },
            required: ['values'],
            additionalProperties: false
          }
        }),
        opt('title', s.STRING),
        opt('stacked', s.STACKED),
        opt('legend', s.LEGEND),
        opt('cell', s.CELL),
        opt('width', s.LENGTH),
        opt('height', s.LENGTH),
        opt('name', s.STRING)
      )
    ]
  },
  removeMacros: { summary: 'Strip VBA so the workbook saves as .xlsx', slots: [] }
}
