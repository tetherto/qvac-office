import type { Package } from './package.ts'

export interface DocumentConversion {
  changed: boolean
  contentType: string
}

const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'
const PML = 'application/vnd.openxmlformats-officedocument.presentationml.'
const SML = 'application/vnd.openxmlformats-officedocument.spreadsheetml.'

const DOCUMENT_TYPES: ReadonlyMap<string, string> = new Map([
  [WML + 'template.main+xml', WML + 'document.main+xml'],
  [
    'application/vnd.ms-word.template.macroEnabledTemplate.main+xml',
    'application/vnd.ms-word.document.macroEnabled.main+xml'
  ],
  [PML + 'template.main+xml', PML + 'presentation.main+xml'],
  [
    'application/vnd.ms-powerpoint.template.macroEnabled.main+xml',
    'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml'
  ],
  [SML + 'template.main+xml', SML + 'sheet.main+xml'],
  [
    'application/vnd.ms-excel.template.macroEnabled.main+xml',
    'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
  ]
])

export function isTemplate(pkg: Package, main: string): boolean {
  return DOCUMENT_TYPES.has(pkg.contentType(main) ?? '')
}

export function toDocument(pkg: Package, main: string): DocumentConversion {
  const current = pkg.contentType(main) ?? ''
  const target = DOCUMENT_TYPES.get(current)
  if (!target) {
    return { changed: false, contentType: current }
  }

  pkg.contentTypes.setOverride(main, target)

  return { changed: true, contentType: target }
}
