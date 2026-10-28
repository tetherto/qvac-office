import { DocsError } from '../errors.ts'
import { Package } from '../opc/package.ts'
import * as pptx from '../pptx/index.ts'
import * as docx from '../docx/index.ts'
import * as xlsx from '../xlsx/index.ts'
import { decode } from './base64.ts'
import { PPTX } from './pptx.ts'
import { DOCX } from './docx.ts'
import { XLSX } from './xlsx.ts'
import {
  opSchema,
  toArguments,
  type Operation,
  type OpTable,
  type ResolveImage,
  type Schema
} from './spec.ts'

export type { Operation, OpSpec, OpTable, Param, Schema, Slot } from './spec.ts'

export type Format = 'pptx' | 'docx' | 'xlsx'
export type Doc = pptx.Presentation | docx.Document | xlsx.Workbook

export interface ToolDefinition {
  name: string
  description: string
  input_schema: Schema
}

export interface OperationDefinition {
  name: string
  description: string
  query: boolean
  input_schema: Schema
}

export interface ToolOptions {
  name?: string
  only?: readonly string[]
}

export interface ApplyOptions {
  resolveImage?: ResolveImage
}

export interface Applied {
  doc: Doc
  results: unknown[]
}

interface Step {
  doc: Doc
  format: Format
  operation: Operation
  resolve: ResolveImage
}

const TABLES: Readonly<Record<Format, OpTable>> = { pptx: PPTX, docx: DOCX, xlsx: XLSX }

const NOUNS: Readonly<Record<Format, string>> = {
  pptx: 'PowerPoint deck',
  docx: 'Word document',
  xlsx: 'Excel workbook'
}

const MAIN_TYPES: readonly [string, Format][] = [
  ['presentationml', 'pptx'],
  ['wordprocessingml', 'docx'],
  ['spreadsheetml', 'xlsx']
]

const OFFICE_DOCUMENT =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'

export function detect(bytes: Uint8Array): Format | null {
  const pkg = Package.open(bytes)
  const main = pkg.related('/', OFFICE_DOCUMENT)[0]
  const type = main ? (pkg.contentType(main) ?? '') : ''
  return MAIN_TYPES.find(([marker]) => type.includes(marker))?.[1] ?? null
}

export function openAs(format: Format, bytes: Uint8Array): Doc {
  if (format === 'pptx') {
    return pptx.open(bytes)
  }

  return format === 'docx' ? docx.open(bytes) : xlsx.open(bytes)
}

export function open(bytes: Uint8Array): Doc {
  const format = detect(bytes)
  if (!format) {
    throw new DocsError('INVALID_INPUT', 'the file is not a PowerPoint, Word or Excel document')
  }

  return openAs(format, bytes)
}

export function formatOf(doc: Doc): Format {
  if (doc instanceof pptx.Presentation) {
    return 'pptx'
  }

  return doc instanceof docx.Document ? 'docx' : 'xlsx'
}

function table(format: Format, only?: readonly string[]): [string, OpTable[string]][] {
  const entries = Object.entries(TABLES[format])
  if (!only) {
    return entries
  }

  const unknown = only.filter((name) => !(name in TABLES[format]))
  if (unknown.length) {
    throw new DocsError('INVALID_OPTIONS', `${format} has no operation ${unknown.join(', ')}`, {
      field: 'only',
      got: unknown,
      candidates: Object.keys(TABLES[format])
    })
  }

  return entries.filter(([name]) => only.includes(name))
}

export function operations(format: Format, opts: ToolOptions = {}): OperationDefinition[] {
  return table(format, opts.only).map(([name, spec]) => ({
    name,
    description: spec.summary,
    query: spec.query === true,
    input_schema: opSchema(name, spec)
  }))
}

export function tool(format: Format, opts: ToolOptions = {}): ToolDefinition {
  const ops = table(format, opts.only)
  return {
    name: opts.name ?? format,
    description: `Read and edit a ${NOUNS[format]}. Pass ops, run in order as one atomic batch: if one fails, none apply. Operations: ${ops.map(([name, spec]) => `${name} (${spec.summary})`).join('; ')}.`,
    input_schema: {
      type: 'object',
      properties: {
        ops: {
          type: 'array',
          minItems: 1,
          items: { anyOf: ops.map(([name, spec]) => opSchema(name, spec)) }
        }
      },
      required: ['ops'],
      additionalProperties: false
    }
  }
}

function defaultImage(value: unknown, key: string): Uint8Array {
  if (value instanceof Uint8Array) {
    return value
  }

  if (typeof value === 'string') {
    return decode(value, key)
  }

  throw new DocsError('INVALID_IMAGE', `${key} must be base64, a data: URL or bytes`, {
    field: key
  })
}

function located(err: unknown, index: number, op: string): unknown {
  if (!(err instanceof DocsError)) {
    return err
  }

  return new DocsError(err.code, `ops[${index}] ${op}: ${err.message}`, {
    ...err.details,
    index,
    op
  })
}

function run({ doc, format, operation, resolve }: Step): unknown {
  const spec = TABLES[format][operation.op]
  if (!spec) {
    throw new DocsError('UNKNOWN_OPERATION', `${format} has no operation ${String(operation.op)}`, {
      field: 'op',
      got: operation.op,
      candidates: Object.keys(TABLES[format])
    })
  }

  const args = toArguments(spec, operation, resolve)
  const method: unknown = Reflect.get(doc, operation.op)
  if (typeof method !== 'function') {
    throw new DocsError('UNKNOWN_OPERATION', `${format} has no method ${operation.op}`)
  }

  return Reflect.apply(method, doc, args)
}

function isList(ops: Operation | readonly Operation[]): ops is readonly Operation[] {
  return Array.isArray(ops)
}

export function apply(
  doc: Doc,
  ops: Operation | readonly Operation[],
  opts: ApplyOptions = {}
): Applied {
  const format = formatOf(doc)
  const resolve = opts.resolveImage ?? defaultImage
  const work = openAs(format, doc.package.save())
  const list = isList(ops) ? ops : [ops]
  const results = list.map((operation, index) => {
    try {
      return run({ doc: work, format, operation, resolve })
    } catch (err) {
      throw located(err, index, String(operation?.op))
    }
  })

  return { doc: work, results }
}
