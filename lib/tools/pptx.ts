import * as s from './schemas.ts'
import { group, opt, renamed, req, type OpTable } from './spec.ts'

const TARGET = group(req('slide', s.SLIDE), req('shape', s.SHAPE))
const BOX = [
  opt('x', s.LENGTH),
  opt('y', s.LENGTH),
  opt('width', s.LENGTH),
  opt('height', s.LENGTH)
]
const CHART = [
  req('type', s.CHART_TYPE),
  req('categories', s.CATEGORIES),
  req('series', s.SERIES),
  opt('title', s.STRING),
  opt('stacked', s.STACKED),
  opt('legend', s.LEGEND),
  opt('name', s.STRING)
]
const INTO = [opt('into', s.SHAPE), opt('replace', s.BOOLEAN)]
const IMAGE = { key: 'image', schema: s.IMAGE, required: true, image: true }

export const PPTX: OpTable = {
  read: {
    summary: 'Slides with their shapes, text, tables, charts and notes',
    slots: [
      group(
        opt('from', s.INDEX),
        opt('count', s.POSITIVE),
        opt('slides', { type: 'array', items: s.SLIDE }),
        opt('detail', { enum: ['runs'] })
      )
    ],
    query: true
  },
  layouts: { summary: 'The layouts a new slide can use', slots: [], query: true },
  outline: {
    summary: 'Each slide in a line: id, layout, title and what it holds',
    slots: [],
    query: true
  },
  readPage: {
    summary: 'Slides within a character budget, with next to continue from',
    slots: [
      group(opt('from', s.SLIDE), opt('limit', s.POSITIVE), opt('detail', { enum: ['runs'] }))
    ],
    query: true
  },
  find: {
    summary: 'Shapes whose text, name, kind or placeholder match',
    slots: [
      group(
        opt('text', s.STRING),
        opt('placeholder', s.STRING),
        opt('kind', s.STRING),
        opt('name', s.STRING),
        opt('regex', s.BOOLEAN),
        opt('matchCase', s.BOOLEAN)
      )
    ],
    query: true
  },
  replaceText: {
    summary: 'Find and replace text across runs, keeping formatting',
    slots: [
      group(
        req('find', s.STRING),
        req('replace', s.STRING),
        opt('all', s.BOOLEAN),
        opt('matchCase', s.BOOLEAN),
        opt('slide', s.SLIDE),
        opt('shape', s.SHAPE),
        opt('notes', s.BOOLEAN)
      )
    ]
  },
  formatText: {
    summary: 'Style only the matched text in a shape',
    slots: [
      TARGET,
      group(req('find', s.STRING), opt('all', s.BOOLEAN), opt('matchCase', s.BOOLEAN)),
      req('style', s.STYLE)
    ]
  },
  link: {
    summary: 'Link the matched text in a shape to an http, https or mailto URL',
    slots: [
      TARGET,
      group(req('find', s.STRING), opt('all', s.BOOLEAN), opt('matchCase', s.BOOLEAN)),
      req('url', s.STRING)
    ]
  },
  setText: {
    summary: "Replace a shape's text, keeping run and paragraph formatting",
    slots: [TARGET, req('text', s.TEXT), group(opt('fit', { enum: ['shrink'] }))]
  },
  setCellText: {
    summary: 'Replace the text of one table cell',
    slots: [
      group(
        req('slide', s.SLIDE),
        req('shape', s.SHAPE),
        req('row', s.INDEX),
        req('column', s.INDEX)
      ),
      req('text', s.TEXT)
    ]
  },
  setStyle: {
    summary: 'Style every run of a shape, or of one paragraph',
    slots: [TARGET, req('style', s.STYLE), group(opt('paragraph', s.INDEX))]
  },
  setNotes: {
    summary: "Replace a slide's speaker notes",
    slots: [req('slide', s.SLIDE), req('text', s.TEXT)]
  },
  addSlide: {
    summary: 'Add a slide from a layout, or like an existing slide',
    slots: [
      group(
        opt('like', s.SLIDE),
        opt('layout', s.STRING),
        opt('at', s.INDEX),
        opt('title', s.TEXT),
        opt('body', s.TEXT),
        { key: 'image', schema: s.IMAGE, image: true }
      )
    ]
  },
  deleteSlide: { summary: 'Delete a slide', slots: [req('slide', s.SLIDE)] },
  moveSlide: {
    summary: 'Move a slide to another position',
    slots: [req('slide', s.SLIDE), group(req('to', s.INDEX))]
  },
  setBackground: {
    summary: "Set a slide's background colour, or clear it with null",
    slots: [req('slide', s.SLIDE), req('color', s.NULLABLE_COLOR)]
  },
  addTextBox: {
    summary: 'Add a text box',
    slots: [
      req('slide', s.SLIDE),
      group(
        ...BOX,
        opt('text', s.TEXT),
        opt('style', s.STYLE),
        opt('fill', s.COLOR),
        opt('name', s.STRING)
      )
    ]
  },
  addTable: {
    summary: 'Add a table',
    slots: [
      req('slide', s.SLIDE),
      group(...BOX, ...INTO, req('rows', s.ROWS), opt('header', s.BOOLEAN), opt('name', s.STRING))
    ]
  },
  setBox: { summary: 'Move or resize a shape', slots: [TARGET, group(...BOX)] },
  deleteShape: { summary: 'Delete a shape', slots: [TARGET] },
  addPicture: {
    summary: 'Add a picture',
    slots: [
      req('slide', s.SLIDE),
      IMAGE,
      group(...BOX, ...INTO, opt('name', s.STRING), opt('description', s.STRING))
    ]
  },
  replacePicture: {
    summary: "Swap a picture's image, keeping its frame",
    slots: [TARGET, IMAGE, group(opt('stretch', s.BOOLEAN))]
  },
  addChart: {
    summary: 'Add a chart',
    slots: [req('slide', s.SLIDE), group(...BOX, ...INTO, ...CHART)]
  },
  setChartData: {
    summary: "Replace a chart's categories and series",
    slots: [TARGET, group(req('categories', s.CATEGORIES), req('series', s.SERIES))]
  },
  setChartTitle: {
    summary: "Set a chart's title, or clear it with null",
    slots: [TARGET, req('title', s.NULLABLE_TEXT)]
  },
  size: { summary: 'The slide size, and its preset when it has one', slots: [], query: true },
  setSize: {
    summary: 'Set the slide size; scale stretches existing shapes to fit',
    slots: [
      req('size', {
        anyOf: [
          { enum: ['16:9', '4:3'] },
          {
            type: 'object',
            properties: { width: s.LENGTH, height: s.LENGTH },
            required: ['width', 'height']
          }
        ]
      }),
      group(opt('scale', s.BOOLEAN))
    ]
  },
  setFill: {
    summary: "Set a shape's solid fill, or remove it with null",
    slots: [TARGET, req('color', s.NULLABLE_COLOR)]
  },
  setLine: {
    summary: "Set a shape's outline, or remove it with null",
    slots: [
      TARGET,
      req('line', {
        anyOf: [
          {
            type: 'object',
            properties: {
              color: s.COLOR,
              width: s.LENGTH,
              dash: {
                enum: [
                  'solid',
                  'dot',
                  'dash',
                  'lgDash',
                  'dashDot',
                  'lgDashDot',
                  'lgDashDotDot',
                  'sysDash',
                  'sysDot',
                  'sysDashDot',
                  'sysDashDotDot'
                ]
              }
            },
            additionalProperties: false
          },
          { type: 'null' }
        ]
      })
    ]
  },
  duplicateShape: {
    summary: 'Copy a shape to the same or another slide',
    slots: [
      TARGET,
      group(renamed(opt('to', s.SLIDE), 'slide'), opt('x', s.LENGTH), opt('y', s.LENGTH))
    ]
  },
  reorder: {
    summary: 'Bring a shape to the front or send it back',
    slots: [TARGET, req('order', { enum: ['front', 'back', 'forward', 'backward'] })]
  },
  group: {
    summary: 'Group shapes on a slide',
    slots: [
      req('slide', s.SLIDE),
      req('shapes', { type: 'array', items: { type: 'integer' }, minItems: 2 }),
      group(opt('name', s.STRING))
    ]
  },
  ungroup: { summary: 'Ungroup a group into its shapes', slots: [TARGET] },
  theme: { summary: 'Theme colours and fonts', slots: [], query: true },
  setTheme: {
    summary: 'Change theme colours and fonts on every master',
    slots: [
      group(
        opt('colors', { type: 'object', additionalProperties: s.COLOR }),
        opt('fonts', {
          type: 'object',
          properties: { heading: s.STRING, body: s.STRING },
          additionalProperties: false
        })
      )
    ]
  },
  setLayout: {
    summary: 'Switch a slide to another layout, carrying placeholder text across',
    slots: [req('slide', s.SLIDE), req('layout', s.STRING)]
  },
  toDocument: { summary: 'Turn a template (.potx) into a regular deck', slots: [] }
}
