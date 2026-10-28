import * as xml from '../xml.ts'
import * as text from './text.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import { W, REL_BASE, WML, blocks, headingLevel, part, styleId, type Host } from './host.ts'
import type { Document, Element } from '../xml.ts'

export interface TocOptions {
  levels?: number
  title?: string
}

interface Entry {
  level: number
  title: string
}

const XML_NS = 'http://www.w3.org/XML/1998/namespace'
const PLACEHOLDER = 'Update the table of contents to list the headings.'
const SETTINGS_ORDER = [
  'writeProtection',
  'view',
  'zoom',
  'removePersonalInformation',
  'removeDateAndTime',
  'doNotDisplayPageBoundaries',
  'displayBackgroundShape',
  'printPostScriptOverText',
  'printFractionalCharacterWidth',
  'printFormsData',
  'embedTrueTypeFonts',
  'embedSystemFonts',
  'saveSubsetFonts',
  'saveFormsData',
  'mirrorMargins',
  'alignBordersAndEdges',
  'bordersDoNotSurroundHeader',
  'bordersDoNotSurroundFooter',
  'gutterAtTop',
  'hideSpellingErrors',
  'hideGrammaticalErrors',
  'activeWritingStyle',
  'proofState',
  'formsDesign',
  'attachedTemplate',
  'linkStyles',
  'stylePaneFormatFilter',
  'stylePaneSortMethod',
  'documentType',
  'mailMerge',
  'revisionView',
  'trackRevisions',
  'doNotTrackMoves',
  'doNotTrackFormatting',
  'documentProtection',
  'autoFormatOverride',
  'styleLockTheme',
  'styleLockQFSet',
  'defaultTabStop',
  'autoHyphenation',
  'consecutiveHyphenLimit',
  'hyphenationZone',
  'doNotHyphenateCaps',
  'showEnvelope',
  'summaryLength',
  'clickAndTypeStyle',
  'defaultTableStyle',
  'evenAndOddHeaders',
  'bookFoldRevPrinting',
  'bookFoldPrinting',
  'bookFoldPrintingSheets',
  'drawingGridHorizontalSpacing',
  'drawingGridVerticalSpacing',
  'displayHorizontalDrawingGridEvery',
  'displayVerticalDrawingGridEvery',
  'doNotUseMarginsForDrawingGridOrigin',
  'drawingGridHorizontalOrigin',
  'drawingGridVerticalOrigin',
  'doNotShadeFormData',
  'noPunctuationKerning',
  'characterSpacingControl',
  'printTwoOnOne',
  'strictFirstAndLastChars',
  'noLineBreaksAfter',
  'noLineBreaksBefore',
  'savePreviewPicture',
  'doNotValidateAgainstSchema',
  'saveInvalidXml',
  'ignoreMixedContent',
  'alwaysShowPlaceholderText',
  'doNotDemarcateInvalidXml',
  'saveXmlDataOnly',
  'useXSLTWhenSaving',
  'saveThroughXslt',
  'showXMLTags',
  'alwaysMergeEmptyNamespace',
  'updateFields',
  'hdrShapeDefaults',
  'footnotePr',
  'endnotePr',
  'compat',
  'docVars',
  'rsids',
  'mathPr',
  'attachedSchema',
  'themeFontLang',
  'clrSchemeMapping',
  'doNotIncludeSubdocsInStats',
  'doNotAutoCompressPictures',
  'forceUpgrade',
  'captions',
  'readModeInkLockDown',
  'smartTagType',
  'schemaLibrary',
  'shapeDefaults',
  'doNotEmbedSmartTags',
  'decimalSymbol',
  'listSeparator'
]

function levelsOf(input: unknown): number {
  const levels = input ?? 3
  if (typeof levels === 'number' && Number.isInteger(levels) && levels >= 1 && levels <= 9) {
    return levels
  }

  throw new DocsError(
    'INVALID_OPTIONS',
    `levels must be an integer from 1 to 9, got ${String(input)}`,
    {
      field: 'levels',
      expected: '1 to 9',
      got: input
    }
  )
}

function entries(host: Host, levels: number): Entry[] {
  const out: Entry[] = []
  for (const el of blocks(host)) {
    const level = headingLevel(host, el)
    const title = el.localName === 'p' ? text.paragraphText(el).trim() : ''
    if (level !== null && level >= 1 && level <= levels && title) {
      out.push({ level, title })
    }
  }

  return out
}

function fieldRun(doc: Document, type: string): Element {
  const r = text.wEl(doc, 'r')
  r.appendChild(
    text.wEl(doc, 'fldChar', { fldCharType: type, dirty: type === 'begin' ? 'true' : undefined })
  )

  return r
}

function instruction(doc: Document, levels: number): Element {
  const r = text.wEl(doc, 'r')
  const instr = xml.append(r, text.wEl(doc, 'instrText'))
  instr.setAttributeNS(XML_NS, 'xml:space', 'preserve')
  instr.appendChild(doc.createTextNode(` TOC \\o "1-${levels}" \\h \\z \\u `))
  return r
}

function entryParagraph(host: Host, doc: Document, entry: Entry): Element {
  const style = styleId(host, `toc ${entry.level}`)
  const p = text.buildParagraph(doc, { styleId: style ?? undefined, text: entry.title }, null)
  if (!style) {
    xml.insertOrdered(
      text.ensure(p, 'pPr'),
      text.wEl(doc, 'ind', { left: 220 * (entry.level - 1) }),
      text.PARAGRAPH_ORDER
    )
  }

  return p
}

function titleParagraph(host: Host, doc: Document, title: string): Element {
  const style = styleId(host, 'TOC Heading')
  const p = style
    ? text.buildParagraph(doc, { styleId: style, text: title }, null)
    : text.buildParagraph(doc, { runs: [{ text: title, bold: true, size: 14 }] }, null)
  xml.insertOrdered(
    text.ensure(p, 'pPr'),
    text.wEl(doc, 'outlineLvl', { val: 9 }),
    text.PARAGRAPH_ORDER
  )

  return p
}

export function build(host: Host, opts: TocOptions = {}): Element[] {
  if (!isRecord(opts)) {
    throw new DocsError('INVALID_OPTIONS', 'give the table of contents as { levels, title }')
  }

  const levels = levelsOf(opts.levels)
  const doc = host.package.xml(host.main)
  const found = entries(host, levels)
  const paragraphs = found.length
    ? found.map((entry) => entryParagraph(host, doc, entry))
    : [text.buildParagraph(doc, PLACEHOLDER, null)]
  const first = paragraphs[0]
  const start =
    first.firstChild && xml.isElement(first.firstChild, W, 'pPr')
      ? first.firstChild.nextSibling
      : first.firstChild
  for (const r of [fieldRun(doc, 'begin'), instruction(doc, levels), fieldRun(doc, 'separate')]) {
    first.insertBefore(r, start)
  }

  paragraphs[paragraphs.length - 1].appendChild(fieldRun(doc, 'end'))
  const title = opts.title === undefined ? [] : [titleParagraph(host, doc, String(opts.title))]
  updateOnOpen(host)

  return [...title, ...paragraphs]
}

export function updateOnOpen(host: Host): void {
  const name = part(host, {
    rel: REL_BASE + 'settings',
    type: WML + 'settings+xml',
    name: '/word/settings.xml',
    root: `<w:settings xmlns:w="${W}"/>`
  })
  const root = host.package.root(name)
  const el =
    xml.child(root, W, 'updateFields') ??
    xml.insertOrdered(root, text.wEl(xml.owner(root), 'updateFields'), SETTINGS_ORDER)
  text.wAttr(el, 'val', 'true')
  host.package.touch(name)
}
