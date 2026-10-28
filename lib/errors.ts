import { CODES, type DocsErrorCode } from './codes.ts'

export class DocsError extends Error {
  code: DocsErrorCode

  constructor(code: DocsErrorCode, message: string) {
    super(message)
    this.name = 'DocsError'
    this.code = code
  }

  static get codes(): typeof CODES {
    return CODES
  }
}
