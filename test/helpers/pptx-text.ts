import * as pptx from '../../lib/pptx/index.ts'
import * as utf8 from '../../lib/utf8.ts'
import { fixture, decode, unzip, zipFiles } from './engine.ts'

export const SLIDE2 = 'ppt/slides/slide2.xml'
export const RICH =
  '<a:p><a:r><a:rPr lang="en-US" b="1" sz="2400"><a:solidFill><a:srgbClr val="1a73e8"/></a:solidFill><a:latin typeface="Georgia"/></a:rPr><a:t>Sun</a:t></a:r><a:r><a:rPr lang="en-US" i="1" u="sng"/><a:t>light</a:t></a:r><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>2</a:t></a:fld><a:br><a:rPr lang="en-US"/></a:br><a:r><a:t>next</a:t></a:r></a:p>'
const ORIGINAL =
  '<a:p><a:r><a:rPr b="1" sz="2400"/><a:t>Sunlight holds every colour</a:t></a:r></a:p>'

export function withParagraph(paragraph: string): pptx.Presentation {
  const files = unzip(fixture('basic.pptx'))
  files[SLIDE2] = utf8.encode(decode(files[SLIDE2]).replace(ORIGINAL, paragraph))
  return pptx.open(zipFiles(files))
}

export function rich(): pptx.Presentation {
  return withParagraph(RICH)
}
