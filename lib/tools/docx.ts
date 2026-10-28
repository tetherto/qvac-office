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
    level: { type: 'integer', minimum: 0, maximum: 8 },
    list: { enum: ['bullet', 'number'] }
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
const SEARCH = group(req('find', s.STRING), opt('all', s.BOOLEAN), opt('matchCase', s.BOOLEAN))
const HEADER = group(opt('type', { enum: ['default', 'first'] }))

export const DOCX: OpTable = {
  read: {
    summary: 'Blocks, headers, footers, footnotes and comments',
    slots: [
      group(
        opt('from', s.INDEX),
        opt('count', s.POSITIVE),
        opt('maxChars', s.POSITIVE),
        opt('detail', { enum: ['runs'] })
      )
    ],
    query: true
  },
  outline: {
    summary: 'Headings with their level and the blocks each spans',
    slots: [group(opt('depth', { type: 'integer', minimum: 1, maximum: 9 }))],
    query: true
  },
  find: {
    summary: 'Blocks whose text, style or heading level match',
    slots: [
      group(
        opt('text', s.STRING),
        opt('regex', s.STRING),
        opt('style', s.STRING),
        opt('heading', { anyOf: [s.BOOLEAN, { type: 'integer', minimum: 1, maximum: 9 }] }),
        opt('matchCase', s.BOOLEAN)
      )
    ],
    query: true
  },
  markdown: {
    summary: 'The body as markdown',
    slots: [group(opt('from', s.INDEX), opt('count', s.POSITIVE))],
    query: true
  },
  page: { summary: 'Page size, orientation and margins', slots: [], query: true },
  sections: { summary: 'Headings and the blocks each one spans', slots: [], query: true },
  paragraphStyles: { summary: 'Paragraph style names', slots: [], query: true },
  setText: {
    summary: "Replace a paragraph's text, keeping its formatting",
    slots: [req('index', s.BLOCK), req('text', TEXT_SPEC)]
  },
  insertParagraphs: {
    summary: 'Insert paragraphs before a block',
    slots: [req('at', s.BLOCK), req('paragraphs', PARAGRAPHS)]
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
      req('at', s.BLOCK),
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
      req('at', s.BLOCK),
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
    slots: [req('at', s.BLOCK), group(opt('width', s.LENGTH), opt('height', s.LENGTH), ...CHART)]
  },
  setChartData: {
    summary: "Replace a chart's categories and series",
    slots: [
      req('index', s.BLOCK),
      group(req('categories', s.CATEGORIES), req('series', s.SERIES)),
      group(opt('chart', s.INDEX))
    ]
  },
  replaceText: {
    summary: 'Find and replace text across runs, keeping formatting',
    slots: [
      group(
        req('find', s.STRING),
        req('replace', s.STRING),
        opt('all', s.BOOLEAN),
        opt('matchCase', s.BOOLEAN),
        opt('in', s.BLOCK)
      )
    ]
  },
  formatText: {
    summary: 'Format only the matched text in a block',
    slots: [req('index', s.BLOCK), SEARCH, req('format', FORMAT)]
  },
  link: {
    summary: 'Link the matched text in a block to an http, https or mailto URL',
    slots: [req('index', s.BLOCK), SEARCH, req('url', s.STRING)]
  },
  appendMarkdown: {
    summary: 'Append markdown: headings, lists, bold, italic, code, links, tables',
    slots: [req('markdown', s.STRING)]
  },
  insertMarkdown: {
    summary: 'Insert markdown before a block',
    slots: [req('at', s.BLOCK), req('markdown', s.STRING)]
  },
  addComment: {
    summary: 'Comment on a paragraph, or on the text found in it',
    slots: [
      req('index', s.BLOCK),
      group(
        req('text', s.STRING),
        opt('author', s.STRING),
        opt('initials', s.STRING),
        opt('find', s.STRING)
      )
    ]
  },
  deleteComment: {
    summary: 'Delete a comment',
    slots: [req('id', { type: ['string', 'integer'] })]
  },
  resolveComment: {
    summary: 'Mark a comment done, or open it again with done false',
    slots: [req('id', { type: ['string', 'integer'] }), group(opt('done', s.BOOLEAN))]
  },
  addFootnote: {
    summary: 'Add a footnote after the found text, or at the end of the paragraph',
    slots: [req('index', s.BLOCK), group(req('text', s.STRING), opt('find', s.STRING))]
  },
  setPage: {
    summary: 'Set page size, orientation and margins',
    slots: [
      group(
        opt('size', {
          anyOf: [
            { enum: ['letter', 'a4'] },
            {
              type: 'object',
              properties: { width: s.LENGTH, height: s.LENGTH },
              required: ['width', 'height']
            }
          ]
        }),
        opt('orientation', { enum: ['portrait', 'landscape'] }),
        opt('margins', {
          type: 'object',
          properties: {
            top: s.LENGTH,
            right: s.LENGTH,
            bottom: s.LENGTH,
            left: s.LENGTH,
            header: s.LENGTH,
            footer: s.LENGTH
          },
          additionalProperties: false
        }),
        opt('all', s.BOOLEAN)
      )
    ]
  },
  addTableOfContents: {
    summary: 'Insert a table of contents that Word refreshes on open',
    slots: [
      req('at', s.BLOCK),
      group(opt('levels', { type: 'integer', minimum: 1, maximum: 9 }), opt('title', s.STRING))
    ]
  },
  acceptChanges: {
    summary: 'Accept tracked changes in the body or one block',
    slots: [group(opt('in', s.BLOCK))]
  },
  rejectChanges: {
    summary: 'Reject tracked changes in the body or one block',
    slots: [group(opt('in', s.BLOCK))]
  },
  toDocument: { summary: 'Turn a template (.dotx) into a regular document', slots: [] },
  setChartTitle: {
    summary: "Set a chart's title, or clear it with null",
    slots: [req('index', s.BLOCK), req('title', s.NULLABLE_TEXT), group(opt('chart', s.INDEX))]
  }
}
