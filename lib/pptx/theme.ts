import * as xml from '../xml.ts'
import * as slides from './slides.ts'
import { DocsError } from '../errors.ts'
import { isRecord } from '../guards.ts'
import { SCHEME_NAMES, type ThemeColor } from './template.ts'
import type { Element } from '../xml.ts'
import type { Theme, ThemeChanged } from './layout-types.ts'

interface Edits {
  colors: [ThemeColor, string][]
  fonts: ['heading' | 'body', string][]
}

const { a: A } = xml.ns
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
const FONT_SLOTS = { heading: 'majorFont', body: 'minorFont' } as const
const THEME_COLORS = Object.keys(SCHEME_NAMES)

function isThemeColor(key: string): key is ThemeColor {
  return key in SCHEME_NAMES
}

export function themeParts(deck: slides.Deck): string[] {
  const masters = deck.pkg.related(deck.main, REL_BASE + 'slideMaster')
  const parts = masters.flatMap((master) => deck.pkg.related(master, REL_BASE + 'theme'))
  if (!parts.length) {
    throw new DocsError('BROKEN_PPTX', 'the slide master has no theme')
  }

  return [...new Set(parts)]
}

function elements(root: Element): Element {
  const found = xml.descendants(root, A, 'themeElements')[0]
  if (!found) {
    throw new DocsError('BROKEN_PPTX', 'the theme has no theme elements')
  }

  return found
}

function colorValue(slot: Element | null): string {
  const color = slot ? xml.children(slot, A)[0] : undefined
  if (!color) {
    return ''
  }

  return color.getAttribute(color.localName === 'sysClr' ? 'lastClr' : 'val') ?? ''
}

function typeface(scheme: Element | null, slot: string): string {
  const font = scheme ? xml.child(scheme, A, slot) : null
  const latin = font ? xml.child(font, A, 'latin') : null
  return latin?.getAttribute('typeface') ?? ''
}

export function read(deck: slides.Deck): Theme {
  const root = deck.pkg.root(themeParts(deck)[0])
  const els = elements(root)
  const scheme = xml.child(els, A, 'clrScheme')
  const fonts = xml.child(els, A, 'fontScheme')
  const colors = { ...SCHEME_NAMES }
  for (const key of Object.keys(colors)) {
    if (isThemeColor(key)) {
      colors[key] = colorValue(
        scheme ? xml.child(scheme, A, SCHEME_NAMES[key]) : null
      ).toUpperCase()
    }
  }

  return {
    name: root.getAttribute('name') ?? '',
    colors,
    fonts: { heading: typeface(fonts, FONT_SLOTS.heading), body: typeface(fonts, FONT_SLOTS.body) }
  }
}

function colorEdits(value: unknown): [ThemeColor, string][] {
  return Object.entries(isRecord(value) ? value : {}).map(([key, color]) => {
    if (!isThemeColor(key)) {
      throw new DocsError(
        'INVALID_OPTIONS',
        `unknown theme color "${key}"; use ${THEME_COLORS.join(', ')}`,
        {
          field: `colors.${key}`,
          candidates: THEME_COLORS
        }
      )
    }

    if (typeof color !== 'string' || !/^#?[0-9a-fA-F]{6}$/.test(color)) {
      throw new DocsError(
        'INVALID_OPTIONS',
        `theme color ${key} must be hex RGB, got ${String(color)}`,
        {
          field: `colors.${key}`,
          expected: 'hex RGB such as "1A73E8"',
          got: color
        }
      )
    }

    return [key, color.replace('#', '').toUpperCase()]
  })
}

function fontEdits(value: unknown): ['heading' | 'body', string][] {
  return Object.entries(isRecord(value) ? value : {}).map(([key, face]) => {
    if (key !== 'heading' && key !== 'body') {
      throw new DocsError('INVALID_OPTIONS', `unknown theme font "${key}"; use heading or body`, {
        field: `fonts.${key}`,
        candidates: ['heading', 'body']
      })
    }

    if (typeof face !== 'string' || !face.trim()) {
      throw new DocsError('INVALID_OPTIONS', `theme font ${key} must be a typeface name`, {
        field: `fonts.${key}`,
        got: face
      })
    }

    return [key, face]
  })
}

function edits(input: unknown): Edits {
  const opts = isRecord(input) ? input : {}
  const out = { colors: colorEdits(opts.colors), fonts: fontEdits(opts.fonts) }
  if (!out.colors.length && !out.fonts.length) {
    throw new DocsError('INVALID_OPTIONS', 'setTheme needs colors or fonts to change')
  }

  return out
}

function child(parent: Element, local: string): Element {
  return xml.child(parent, A, local) ?? xml.append(parent, xml.builder(xml.owner(parent), A)(local))
}

function writeColor(scheme: Element, key: ThemeColor, color: string): void {
  const slot = child(scheme, SCHEME_NAMES[key])
  for (const el of xml.children(slot)) {
    xml.remove(el)
  }

  slot.appendChild(xml.builder(xml.owner(slot), A)('srgbClr', { val: color }))
}

function writeFont(scheme: Element, key: 'heading' | 'body', face: string): void {
  const font = child(scheme, FONT_SLOTS[key])
  const a = xml.builder(xml.owner(font), A)
  let latin = xml.child(font, A, 'latin')
  if (!latin) {
    latin = a('latin')
    font.insertBefore(latin, font.firstChild)
  }

  latin.setAttribute('typeface', face)
}

export function write(deck: slides.Deck, input: unknown): ThemeChanged {
  const { colors, fonts } = edits(input)
  const parts = themeParts(deck)
  for (const part of parts) {
    const els = elements(deck.pkg.root(part))
    for (const [key, color] of colors) {
      writeColor(child(els, 'clrScheme'), key, color)
    }

    for (const [key, face] of fonts) {
      writeFont(child(els, 'fontScheme'), key, face)
    }

    deck.pkg.touch(part)
  }

  return {
    changed: [...colors.map(([key]) => key), ...fonts.map(([key]) => key)],
    parts: parts.length
  }
}
