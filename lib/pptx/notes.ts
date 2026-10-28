import * as xml from '../xml.ts'
import * as text from '../drawingml/text.ts'
import type { Document, Element } from '../xml.ts'
import { SHAPE_ORDER, shapesOf, nextShapeId } from './shapes.ts'
import { notesTheme } from './template.ts'
import { PRESENTATION_ORDER } from './slides.ts'
import type { Deck } from './slides.ts'

const { a: A, p: P, r: R } = xml.ns
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
const REL = {
  slide: REL_BASE + 'slide',
  notesSlide: REL_BASE + 'notesSlide',
  notesMaster: REL_BASE + 'notesMaster',
  theme: REL_BASE + 'theme'
}
const NOTES_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'
const NOTES_MASTER_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml'
const THEME_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml'

const NOTES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:notes xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`

const NOTES_MASTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:notesMaster xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="381000" y="685800"/><a:ext cx="6096000" cy="3429000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4343400"/><a:ext cx="5486400" cy="4114800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:notesStyle></p:notesMaster>`

export function notesBody(doc: Document, id: number): Element {
  const p = xml.builder(doc, P)
  const sp = p('sp')
  const nv = xml.append(sp, p('nvSpPr'))
  nv.appendChild(p('cNvPr', { id, name: `Notes Placeholder ${id - 1}` }))
  xml.append(nv, p('cNvSpPr')).appendChild(xml.builder(doc, A)('spLocks', { noGrp: 1 }))
  xml.append(nv, p('nvPr')).appendChild(p('ph', { type: 'body', idx: 1 }))
  sp.appendChild(p('spPr'))
  text.ensureTextBody(sp, SHAPE_ORDER)
  return sp
}

function createNotesMaster(deck: Deck): string {
  const { pkg } = deck
  const themeSource = pkg.related(deck.main, REL.theme)[0]
  const theme = pkg.freeName('/ppt/theme/theme%d.xml')
  pkg.add(theme, themeSource ? pkg.bytes(themeSource) : notesTheme(), THEME_TYPE)
  const part = pkg.freeName('/ppt/notesMasters/notesMaster%d.xml')
  pkg.add(part, xml.parse(NOTES_MASTER_XML), NOTES_MASTER_TYPE)
  pkg.rels(part).add(REL.theme, theme)
  const rId = pkg.rels(deck.main).add(REL.notesMaster, part)
  const top = pkg.root(deck.main)
  const p = xml.builder(xml.owner(top), P)
  const list =
    xml.child(top, P, 'notesMasterIdLst') ??
    xml.insertOrdered(top, p('notesMasterIdLst'), PRESENTATION_ORDER)
  const ref = xml.append(list, p('notesMasterId'))
  xml.setNsAttribute(ref, { namespace: R, local: 'id' }, rId)
  pkg.touch(deck.main)

  return part
}

function createNotes(deck: Deck, slidePart: string): string {
  const { pkg } = deck
  const master = pkg.related(deck.main, REL.notesMaster)[0] ?? createNotesMaster(deck)
  const part = pkg.freeName('/ppt/notesSlides/notesSlide%d.xml')
  const doc = xml.parse(NOTES_XML)
  xml.descendants(doc, P, 'spTree')[0]?.appendChild(notesBody(doc, 3))
  pkg.add(part, doc, NOTES_TYPE)
  pkg.rels(part).add(REL.notesMaster, master)
  pkg.rels(part).add(REL.slide, slidePart)
  pkg.rels(slidePart).add(REL.notesSlide, part)
  return part
}

function bodyOf(doc: Document): Element {
  const existing = shapesOf(doc).find((s) => s.placeholder === 'body')
  if (existing) {
    return existing.el
  }

  const el = notesBody(doc, nextShapeId(doc))
  xml.descendants(doc, P, 'spTree')[0]?.appendChild(el)

  return el
}

export function writeNotes(deck: Deck, slidePart: string, value: unknown): void {
  text.normalizeParagraphs(value)
  const notes = deck.pkg.related(slidePart, REL.notesSlide)[0] ?? createNotes(deck, slidePart)
  const txBody = text.ensureTextBody(bodyOf(deck.pkg.xml(notes)), SHAPE_ORDER)
  text.writeParagraphs(txBody, value)
  deck.pkg.touch(notes)
}

export function notesText(deck: Deck, slidePart: string): string | null {
  const notes = deck.pkg.related(slidePart, REL.notesSlide)[0]
  if (!notes) {
    return null
  }

  const body = shapesOf(deck.pkg.xml(notes)).find((s) => s.placeholder === 'body')
  const txBody = body ? xml.child(body.el, P, 'txBody') : null
  const value = txBody ? text.readText(txBody).trim() : ''

  return value || null
}
