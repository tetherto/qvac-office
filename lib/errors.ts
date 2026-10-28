import { CODES, type DocsErrorCode } from './codes.ts'

export type ErrorDetails = Readonly<Record<string, unknown>>

export class DocsError extends Error {
  code: DocsErrorCode
  details?: ErrorDetails

  constructor(code: DocsErrorCode, message: string, details?: ErrorDetails) {
    super(message)
    this.name = 'DocsError'
    this.code = code
    if (details) {
      this.details = details
    }
  }

  static get codes(): typeof CODES {
    return CODES
  }

  toJSON(): { code: DocsErrorCode; message: string; details?: ErrorDetails } {
    return this.details
      ? { code: this.code, message: this.message, details: this.details }
      : { code: this.code, message: this.message }
  }
}
