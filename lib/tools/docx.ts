import * as s from './schemas.ts'
import { group, opt, req, type OpTable, type Schema } from './spec.ts'

const RUN_FORMAT = {
  bold: s.BOOLEAN,
  italic: s.BOOLEAN,
  underline: s.BOOLEAN,
  size: s.NUMBER,
  color: s.COLOR,
  font: s.STRING
}

const RUN: Schema = {
  type: 'object',
  properties: { text: s.STRING, ...RUN_FORMAT },
  required: ['text'],
  additionalProperties: false
}

const PARAGRAPH_SPEC: Schema = {
  type: 'object',
  properties: {
    text: s.STRING,
    runs: { type: 'array', items: { anyOf: [s.STRING, RUN] } },
    style: { type: 'string', description: 'a style name such as "Heading 1"' },
    styleId: s.STRING,
    align: s.ALIGN,
    pageBreak: s.BOOLEAN,
    level: { type: 'integer', minimum: 0, maximum: 8 }
  },
  additionalProperties: false
}

const TEXT_SPEC: Schema = { anyOf: [s.STRING, PARAGRAPH_SPEC] }
const PARAGRAPHS: Schema = { anyOf: [TEXT_SPEC, { type: 'array', items: TEXT_SPEC }] }
const NULLABLE_PARAGRAPHS: Schema = { anyOf: [PARAGRAPHS, { type: 'null' }] }
const FORMAT: Schema = {
  type: 'object',
  properties: { ...RUN_FORMAT, align: s.ALIGN },
  additionalProperties: false
}
const CHART = [
  req('type', s.CHART_TYPE),
  req('categories', s.CATEGORIES),
  req('series', s.SERIES),
  opt('title', s.STRING),
  opt('stacked', s.STACKED),
  opt('legend', s.LEGEND),
  opt('name', s.STRING)
]
const HEADER = group(opt('type', { enum: ['default', 'first'] }))

export const DOCX: OpTable = {
  read: {
    summary: 'Blocks, headers, footers, footnotes and comments',
    slots: [group(opt('from', s.INDEX), opt('count', s.POSITIVE))],
    query: true
  },
  sections: { summary: 'Headings and the blocks each one spans', slots: [], query: true },
  paragraphStyles: { summary: 'Paragraph style names', slots: [], query: true },
  setText: {
    summary: "Replace a paragraph's text, keeping its formatting",
    slots: [req('index', s.BLOCK), req('text', TEXT_SPEC)]
  },
  insertParagraphs: {
    summary: 'Insert paragraphs before a block',
    slots: [req('at', s.INDEX), req('paragraphs', PARAGRAPHS)]
  },
  appendParagraphs: {
    summary: 'Append paragraphs at the end',
    slots: [req('paragraphs', PARAGRAPHS)]
  },
  deleteBlocks: {
    summary: 'Delete blocks',
    slots: [req('index', s.BLOCK), group(opt('count', s.POSITIVE))]
  },
  formatParagraph: {
    summary: "Format a paragraph's runs and alignment",
    slots: [req('index', s.BLOCK), req('format', FORMAT)]
  },
  formatStyle: {
    summary: 'Change a named style everywhere it is used',
    slots: [req('name', s.STRING), req('format', FORMAT)]
  },
  replaceSection: {
    summary: "Replace a section's heading and body",
    slots: [req('index', s.INDEX), group(opt('title', TEXT_SPEC), opt('body', PARAGRAPHS))]
  },
  deleteSection: {
    summary: 'Delete a heading and everything under it',
    slots: [req('index', s.INDEX)]
  },
  setHeader: {
    summary: 'Set the page header, or remove it with null',
    slots: [req('text', NULLABLE_PARAGRAPHS), HEADER]
  },
  setFooter: {
    summary: 'Set the page footer, or remove it with null',
    slots: [req('text', NULLABLE_PARAGRAPHS), HEADER]
  },
  addTable: {
    summary: 'Insert a table before a block',
    slots: [
      req('at', s.INDEX),
      group(req('rows', s.ROWS), opt('header', s.BOOLEAN), opt('style', s.STRING))
    ]
  },
  setCell: {
    summary: "Replace a table cell's text",
    slots: [
      group(req('table', s.BLOCK), req('row', s.INDEX), req('column', s.INDEX)),
      req('text', TEXT_SPEC)
    ]
  },
  addRow: {
    summary: 'Add a table row',
    slots: [
      req('table', s.BLOCK),
      req('values', { type: 'array', items: { type: ['string', 'number', 'null'] } }),
      group(opt('at', s.INDEX))
    ]
  },
  deleteRow: {
    summary: 'Delete a table row',
    slots: [group(req('table', s.BLOCK), req('row', s.INDEX))]
  },
  addImage: {
    summary: 'Insert an image before a block',
    slots: [
      req('at', s.INDEX),
      { key: 'image', schema: s.IMAGE, required: true, image: true },
      group(
        opt('width', s.LENGTH),
        opt('height', s.LENGTH),
        opt('align', s.ALIGN),
        opt('name', s.STRING),
        opt('description', s.STRING)
      )
    ]
  },
  addChart: {
    summary: 'Insert a chart before a block',
    slots: [req('at', s.INDEX), group(opt('width', s.LENGTH), opt('height', s.LENGTH), ...CHART)]
  },
  setChartData: {
    summary: "Replace a chart's categories and series",
    slots: [
      req('index', s.BLOCK),
      group(req('categories', s.CATEGORIES), req('series', s.SERIES)),
      group(opt('chart', s.INDEX))
    ]
  },
  setChartTitle: {
    summary: "Set a chart's title, or clear it with null",
    slots: [req('index', s.BLOCK), req('title', s.NULLABLE_TEXT), group(opt('chart', s.INDEX))]
  }
}
