# Corpus

These files come from the test data of [Apache POI](https://github.com/apache/poi)
(`test-data/slideshow`, `test-data/document` and `test-data/spreadsheet` on the
`trunk` branch), which is licensed under the Apache License, Version 2.0. They
were saved by PowerPoint, Word, Excel, LibreOffice and other producers, and
include hostile samples — XML bombs, external entities, Strict Open XML,
truncated archives and fuzzer-minimized crashes — that must fail cleanly.

`expected.json` records the outcome of opening, editing, saving and reopening
each file: `ok`, or the `DocsError` code it must fail with. `run.js` produces
that record; the corpus test compares against it, so a crash, a new failure or
a newly accepted file all show up as a difference to review.

To add a file, copy it into the folder for its format, run
`node test/corpus/run.js > test/corpus/expected.json`, and review the change.

`diff.js` prints only the files whose outcome differs from the record, which is
handy on a runtime whose test runner cannot print a diff.
