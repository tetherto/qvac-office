declare module 'brittle' {
  interface Assertion {
    is(actual: unknown, expected: unknown, message?: string): void
    not(actual: unknown, expected: unknown, message?: string): void
    alike(actual: unknown, expected: unknown, message?: string): void
    unlike(actual: unknown, expected: unknown, message?: string): void
    ok(value: unknown, message?: string): void
    absent(value: unknown, message?: string): void
    pass(message?: string): void
    fail(message?: string): void
    exception(fn: unknown, error?: unknown, message?: string): Promise<void>
    execution(fn: unknown, message?: string): Promise<void>
    comment(message: string): void
    plan(count: number): void
    teardown(fn: () => unknown): void
  }

  type Test = (t: Assertion) => unknown

  interface Brittle {
    (name: string, fn: Test): void
    (name: string, opts: object, fn: Test): void
    skip(name: string, fn: Test): void
    solo(name: string, fn: Test): void
    pause(): void
    resume(): void
    load(url: string): Promise<void>
  }

  const test: Brittle
  export default test
}
