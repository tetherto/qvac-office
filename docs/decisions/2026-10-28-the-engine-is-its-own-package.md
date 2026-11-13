---
id: ADR-2026-10-28-the-engine-is-its-own-package
status: accepted
date: 2026-10-28
---

# The office engine is its own package

## Decision

The engine that lived in the harness at `lib/office/engine/` is published as `@qvac/office`.
It keeps its API, its tests and its corpus unchanged; the harness and any app import it from npm.

## Why

An app needs the engine without the harness: a desktop editor drives it directly to exercise every
edit an LLM can make. `2026-09-28-office-files-are-edited-by-the-harness-engine.md` named this as
the condition to extract it again.

## Consequences

- The package ships compiled `dist/` with declarations; `lib/` is the source.
- Bare and Node both run the test suite on every change, against the TypeScript source.
- The harness's office tools become a consumer of this package's public entries.
