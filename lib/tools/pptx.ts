import * as s from './schemas.ts'
import { group, opt, req, type OpTable } from './spec.ts'

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
const IMAGE = { key: 'image', schema: s.IMAGE, required: true, image: true }

export const PPTX: OpTable = {
  read: {
    summary: 'Slides with their shapes, text, tables, charts and notes',
    slots: [
      group(
        opt('from', s.INDEX),
        opt('count', s.POSITIVE),
        opt('slides', { type: 'array', items: s.INDEX })
      )
    ],
    query: true
  },
  layouts: { summary: 'The layouts a new slide can use', slots: [], query: true },
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
      group(...BOX, req('rows', s.ROWS), opt('header', s.BOOLEAN), opt('name', s.STRING))
    ]
  },
  setBox: { summary: 'Move or resize a shape', slots: [TARGET, group(...BOX)] },
  deleteShape: { summary: 'Delete a shape', slots: [TARGET] },
  addPicture: {
    summary: 'Add a picture',
    slots: [
      req('slide', s.SLIDE),
      IMAGE,
      group(...BOX, opt('name', s.STRING), opt('description', s.STRING))
    ]
  },
  replacePicture: {
    summary: "Swap a picture's image, keeping its frame",
    slots: [TARGET, IMAGE, group(opt('stretch', s.BOOLEAN))]
  },
  addChart: { summary: 'Add a chart', slots: [req('slide', s.SLIDE), group(...BOX, ...CHART)] },
  setChartData: {
    summary: "Replace a chart's categories and series",
    slots: [TARGET, group(req('categories', s.CATEGORIES), req('series', s.SERIES))]
  },
  setChartTitle: {
    summary: "Set a chart's title, or clear it with null",
    slots: [TARGET, req('title', s.NULLABLE_TEXT)]
  }
}
