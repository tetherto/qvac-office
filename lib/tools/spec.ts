import { DocsError } from '../errors.ts'

export type Schema = Readonly<Record<string, unknown>>

export interface Param {
  readonly key: string
  readonly schema: Schema
  readonly required?: boolean
  readonly image?: boolean
  readonly name?: string
}

export interface Group {
  readonly group: readonly Param[]
}

export type Slot = Param | Group

export interface OpSpec {
  readonly summary: string
  readonly slots: readonly Slot[]
  readonly query?: boolean
}

export type OpTable = Readonly<Record<string, OpSpec>>

export interface Operation {
  readonly op: string
  readonly [key: string]: unknown
}

export type ResolveImage = (value: unknown, key: string) => Uint8Array

export function req(key: string, schema: Schema): Param {
  return { key, schema, required: true }
}

export function opt(key: string, schema: Schema): Param {
  return { key, schema }
}

export function renamed(param: Param, name: string): Param {
  return { ...param, name }
}

export function group(...params: Param[]): Group {
  return { group: params }
}

export function params(spec: OpSpec): Param[] {
  return spec.slots.flatMap((slot) => ('group' in slot ? slot.group : [slot]))
}

export function opSchema(name: string, spec: OpSpec): Schema {
  const all = params(spec)
  return {
    type: 'object',
    description: spec.summary,
    properties: {
      op: { const: name },
      ...Object.fromEntries(all.map((param) => [param.key, param.schema]))
    },
    required: ['op', ...all.filter((param) => param.required).map((param) => param.key)],
    additionalProperties: false
  }
}

function checkKeys(name: string, spec: OpSpec, operation: Operation): void {
  const known = new Set(params(spec).map((param) => param.key))
  const unknown = Object.keys(operation).filter((key) => key !== 'op' && !known.has(key))
  if (unknown.length) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `${name} does not take ${unknown.join(', ')}; it takes ${[...known].join(', ') || 'nothing'}`,
      { op: name, field: unknown[0], expected: [...known] }
    )
  }

  const missing = params(spec).filter(
    (param) => param.required && operation[param.key] === undefined
  )
  if (missing.length) {
    throw new DocsError(
      'INVALID_OPTIONS',
      `${name} needs ${missing.map((param) => param.key).join(', ')}`,
      { op: name, field: missing[0]?.key, expected: missing.map((param) => param.key) }
    )
  }
}

function value(param: Param, operation: Operation, resolve: ResolveImage): unknown {
  const raw = operation[param.key]
  return param.image && raw !== undefined ? resolve(raw, param.key) : raw
}

export function toArguments(spec: OpSpec, operation: Operation, resolve: ResolveImage): unknown[] {
  checkKeys(operation.op, spec, operation)
  const args = spec.slots.map((slot) => {
    if (!('group' in slot)) {
      return value(slot, operation, resolve)
    }

    const present = slot.group.filter((param) => operation[param.key] !== undefined)

    return present.length
      ? Object.fromEntries(
          present.map((param) => [param.name ?? param.key, value(param, operation, resolve)])
        )
      : undefined
  })
  while (args.length && args.at(-1) === undefined) {
    args.pop()
  }

  return args
}
