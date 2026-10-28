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
const DIFF_STYLE: Schema = {
  type: 'object',
  properties: {
    bold: s.BOOLEAN,
    italic: s.BOOLEAN,
    underline: s.BOOLEAN,
    color: s.COLOR,
    fill: s.COLOR
  },
  additionalProperties: false
}
const RULE_VALUE: Schema = { type: ['number', 'string'], description: 'a value, or "=FORMULA"' }
const CONDITIONAL_RULE: Schema = {
  anyOf: [
    {
      type: 'object',
      properties: {
        kind: { const: 'cellIs' },
        operator: {
          enum: [
            'equal',
            'notEqual',
            'greaterThan',
            'greaterThanOrEqual',
            'lessThan',
            'lessThanOrEqual',
            'between',
            'notBetween'
          ]
        },
        value: RULE_VALUE,
        values: { type: 'array', items: RULE_VALUE },
        style: DIFF_STYLE
      },
      required: ['kind', 'operator', 'style'],
      additionalProperties: false
    },
    {
      type: 'object',
      properties: {
        kind: { const: 'colorScale' },
        colors: { type: 'array', items: s.COLOR, minItems: 2, maxItems: 3 }
      },
      required: ['kind', 'colors'],
      additionalProperties: false
    },
    {
      type: 'object',
      properties: { kind: { const: 'dataBar' }, color: s.COLOR },
      required: ['kind', 'color'],
      additionalProperties: false
    },
    {
      type: 'object',
      properties: { kind: { const: 'text' }, contains: s.STRING, style: DIFF_STYLE },
      required: ['kind', 'contains', 'style'],
      additionalProperties: false
    }
  ]
}
const MESSAGE: Schema = {
  anyOf: [
    s.STRING,
    {
      type: 'object',
      properties: {
        title: s.STRING,
        text: s.STRING,
        style: { enum: ['stop', 'warning', 'information'] }
      },
      required: ['text']
    }
  ]
}
const ROW_SPAN: Schema = {
  anyOf: [s.POSITIVE, { type: 'array', items: s.POSITIVE, minItems: 2, maxItems: 2 }]
}
const COLUMN_SPAN: Schema = {
  anyOf: [s.COLUMN, { type: 'array', items: s.COLUMN, minItems: 2, maxItems: 2 }]
}

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
  removeMacros: { summary: 'Strip VBA so the workbook saves as .xlsx', slots: [] },
  outline: {
    summary: 'Each sheet in a line: used range, header row guess, sizes, charts and tables',
    slots: [],
    query: true
  },
  find: {
    summary: 'Cells whose value (or formula) matches',
    slots: [
      req('query', s.STRING),
      group(
        opt('sheet', s.SHEET),
        opt('regex', s.BOOLEAN),
        opt('matchCase', s.BOOLEAN),
        opt('formulas', s.BOOLEAN),
        opt('limit', s.POSITIVE)
      )
    ],
    query: true
  },
  insertColumns: {
    summary: 'Insert columns, moving every reference to the right',
    slots: [SHEET, group(req('column', s.COLUMN), opt('count', s.POSITIVE), opt('values', ROWS))]
  },
  merge: {
    summary: 'Merge a range, keeping the top-left value',
    slots: [SHEET, req('range', s.RANGE)]
  },
  unmerge: {
    summary: 'Unmerge every merge the range touches',
    slots: [SHEET, req('range', s.RANGE)]
  },
  setRowHeight: {
    summary: "Set a row's height in points or as a length",
    slots: [SHEET, group(req('row', s.POSITIVE), req('height', { anyOf: [s.NUMBER, s.LENGTH] }))]
  },
  hide: {
    summary: 'Hide or show rows and columns',
    slots: [
      SHEET,
      group(opt('rows', ROW_SPAN), opt('columns', COLUMN_SPAN), opt('hidden', s.BOOLEAN))
    ]
  },
  freeze: {
    summary: 'Freeze the top rows and left columns; 0 and 0 unfreezes',
    slots: [SHEET, group(opt('rows', s.INDEX), opt('columns', s.INDEX))]
  },
  autofilter: {
    summary: 'Set the sheet filter range, or clear it with null',
    slots: [SHEET, req('range', { anyOf: [s.RANGE, { type: 'null' }] })]
  },
  sort: {
    summary: 'Sort the rows of a range by one or more columns',
    slots: [
      SHEET,
      group(
        req('range', s.RANGE),
        req('by', {
          type: 'array',
          items: {
            type: 'object',
            properties: { column: s.COLUMN, descending: s.BOOLEAN },
            required: ['column'],
            additionalProperties: false
          }
        }),
        opt('header', s.BOOLEAN)
      )
    ]
  },
  addConditionalFormat: {
    summary: 'Highlight cells by value, text, colour scale or data bar',
    slots: [SHEET, req('range', s.RANGE), req('rule', CONDITIONAL_RULE)]
  },
  deleteConditionalFormats: {
    summary: 'Delete conditional formats on the sheet, or those touching a range',
    slots: [SHEET, opt('range', s.RANGE)]
  },
  addValidation: {
    summary: 'Restrict what cells accept: a list, numbers, dates or text length',
    slots: [
      SHEET,
      req('range', s.RANGE),
      group(
        req('type', { enum: ['list', 'whole', 'decimal', 'date', 'textLength'] }),
        opt('values', { anyOf: [{ type: 'array', items: s.STRING }, s.STRING] }),
        opt('min', { type: ['number', 'string'] }),
        opt('max', { type: ['number', 'string'] }),
        opt('operator', {
          enum: [
            'between',
            'notBetween',
            'equal',
            'notEqual',
            'greaterThan',
            'greaterThanOrEqual',
            'lessThan',
            'lessThanOrEqual'
          ]
        }),
        opt('allowBlank', s.BOOLEAN),
        opt('prompt', MESSAGE),
        opt('error', MESSAGE)
      )
    ]
  },
  deleteValidations: {
    summary: 'Delete validations on the sheet, or those touching a range',
    slots: [SHEET, opt('range', s.RANGE)]
  },
  names: { summary: 'Defined names with their references and scope', slots: [], query: true },
  defineName: {
    summary: 'Define or replace a named range',
    slots: [req('name', s.STRING), req('ref', s.STRING), group(opt('sheet', s.SHEET))]
  },
  deleteName: {
    summary: 'Delete a defined name',
    slots: [req('name', s.STRING), group(opt('sheet', s.SHEET))]
  },
  tables: {
    summary: 'Tables with their range, columns and style',
    slots: [opt('sheet', s.SHEET)],
    query: true
  },
  addTable: {
    summary: 'Turn a range into a table with a header row',
    slots: [
      SHEET,
      req('range', s.RANGE),
      group(
        opt('name', s.STRING),
        opt('header', s.BOOLEAN),
        opt('style', s.STRING),
        opt('totals', s.BOOLEAN)
      )
    ]
  },
  deleteTable: {
    summary: 'Delete a table, rewriting its structured references as plain ranges',
    slots: [req('name', s.STRING)]
  }
}
