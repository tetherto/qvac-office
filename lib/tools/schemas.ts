import type { Schema } from './spec.ts'

export const STRING: Schema = { type: 'string' }
export const BOOLEAN: Schema = { type: 'boolean' }
export const NUMBER: Schema = { type: 'number' }
export const INDEX: Schema = { type: 'integer', minimum: 0 }
export const POSITIVE: Schema = { type: 'integer', minimum: 1 }

export const LENGTH: Schema = {
  anyOf: [
    { type: 'integer', minimum: 0, description: 'EMU' },
    {
      type: 'string',
      pattern: '^\\s*\\d+(\\.\\d+)?\\s*(emu|in|cm|mm|pt|px)\\s*$',
      description: 'a length such as "2in", "5cm", "40mm", "72pt" or "96px"'
    }
  ]
}

export const COLOR: Schema = {
  type: 'string',
  pattern: '^#?[0-9A-Fa-f]{6}$',
  description: 'RGB hex such as "1F4E79"'
}

export const NULLABLE_COLOR: Schema = { anyOf: [COLOR, { type: 'null' }] }

export const IMAGE: Schema = {
  type: 'string',
  description: 'the image as base64 or a data: URL (PNG, JPEG or GIF)'
}

export const ALIGN: Schema = { enum: ['left', 'center', 'right', 'justify'] }

export const PARAGRAPH: Schema = {
  type: 'object',
  properties: { text: STRING, level: { type: 'integer', minimum: 0, maximum: 8 } },
  required: ['text'],
  additionalProperties: false
}

export const TEXT: Schema = {
  anyOf: [
    { type: 'string', description: 'newlines separate paragraphs' },
    { type: 'array', items: { anyOf: [STRING, PARAGRAPH] } }
  ]
}

export const NULLABLE_TEXT: Schema = { anyOf: [TEXT, { type: 'null' }] }

export const STYLE: Schema = {
  type: 'object',
  properties: {
    size: { type: 'number', description: 'points' },
    bold: BOOLEAN,
    italic: BOOLEAN,
    underline: BOOLEAN,
    color: COLOR,
    font: STRING,
    align: ALIGN
  },
  additionalProperties: false
}

export const ROWS: Schema = {
  type: 'array',
  items: { type: 'array', items: { type: ['string', 'number', 'null'] } }
}

export const CHART_TYPE: Schema = { enum: ['column', 'bar', 'line', 'pie', 'doughnut', 'area'] }
export const LEGEND: Schema = { enum: [false, 'right', 'left', 'top', 'bottom'] }
export const STACKED: Schema = { enum: [true, false, 'percent'] }

export const CATEGORIES: Schema = { type: 'array', items: { type: ['string', 'number'] } }

export const SERIES: Schema = {
  type: 'array',
  items: {
    type: 'object',
    properties: { name: STRING, values: { type: 'array', items: { type: ['number', 'null'] } } },
    required: ['name', 'values'],
    additionalProperties: false
  }
}

export const SLIDE: Schema = {
  anyOf: [
    { type: 'integer', minimum: 0, description: 'slide index from 0' },
    {
      type: 'object',
      properties: { id: { type: 'integer' } },
      required: ['id'],
      additionalProperties: false,
      description: 'stable slide id from read()'
    }
  ]
}

export const SHAPE: Schema = {
  anyOf: [
    { type: 'integer', description: 'shape id from read()' },
    { enum: ['title', 'subtitle', 'body'] }
  ]
}

export const BLOCK: Schema = {
  anyOf: [
    { type: 'integer', minimum: 0, description: 'block index from 0' },
    {
      type: 'object',
      properties: { id: STRING },
      required: ['id'],
      additionalProperties: false,
      description: 'stable block id from read()'
    }
  ]
}

export const SHEET: Schema = {
  anyOf: [STRING, INDEX],
  description: 'sheet name or index from 0'
}

export const CELL: Schema = { type: 'string', pattern: '^\\$?[A-Za-z]{1,3}\\$?\\d+$' }
export const RANGE: Schema = { type: 'string', description: 'an A1 range such as "A1:C10"' }
export const COLUMN: Schema = { anyOf: [STRING, POSITIVE], description: 'letter or number from 1' }

export const CELL_VALUE: Schema = {
  anyOf: [{ type: ['string', 'number', 'boolean', 'null'] }, { type: 'object' }],
  description: 'a value, "=FORMULA", or { value | formula | date, ...style }'
}
