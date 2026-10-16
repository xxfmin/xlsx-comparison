# workbook-parity

Proves that an `.xlsx` produced with Apache POI 5.3.0 is a 1:1 replica of a reference workbook
(e.g. the legacy generator's output): data, styling, layout, and everything else in the package.

## How it works

Each workbook is reduced to a canonical snapshot (`WorkbookSnapshot`), and two snapshots are
compared key by key (`WorkbookComparator`). Every part of the package is covered:

- **Semantically**, where meaning matters: cell values (exact, no epsilon), formulas and cached
  results, resolved styles (theme/indexed/tinted colours resolved to the RGB that renders),
  column widths, row heights, merges, print setup, and so on.
- **As canonical XML** for everything else: attributes sorted, schema defaults dropped, booleans and
  numbers normalised, so encoding-only differences (`"1"` vs `"true"`, style-table order, shared vs
  inline strings, how column ranges are split) never count.
- **By content hash** for any part referenced from the workbook: images, charts, OLE objects,
  header/footer logos, VBA, custom XML.

Results come in two tiers:

| Tier | Meaning | Fails the comparison |
|---|---|---|
| Difference | Changes what a reader sees or how the file behaves | Always |
| Informational | Bookkeeping: application stamps, window geometry, salted password hashes, ... | Only with `--strict` |

## Usage

Requires JDK 25 (the code also compiles on 21) and Maven.

```bash
mvn test                                   # self-tests, gap coverage, render validation
mvn test -Dparity.expected=legacy/ -Dparity.actual=generated/   # parity corpus as JUnit tests

# Batch comparison of two directories (files are paired by name)
mvn -q compile exec:java -Dexec.args="legacy/ generated/ parity-reports/"
mvn -q compile exec:java -Dexec.args="--strict --password=SECRET legacy/ generated/"
```

The batch run exits `0` only when every pair passes; each pair that is not identical gets a full
report in `parity-reports/<file>.txt`. Generate both sides from the same input data with a fixed
clock, so run dates don't show up as differences.

## Tests

| Suite | What it proves |
|---|---|
| `WorkbookComparatorSelfTest` | Real changes are caught; encoding-only changes are not |
| `GapCoverageTest` | Each gap found by walking the ECMA-376 schema (G1-G18) lands in the right tier |
| `RenderEquivalenceTest` | The comparator's "renders identically" claims hold when LibreOffice prints both files (skipped without `soffice` and `pdftoppm`) |
| `StatementParityTest` | Your corpus; skipped unless `-Dparity.expected` is set |

## Known limitations

- Rendering claims are validated with LibreOffice, not Excel.
- A cell with no style index in a styled row/column is reported rather than resolved: the schema
  default is style 0, but POI and LibreOffice apply the row/column style.
- POI's streaming writer (SXSSF) produces a zip container LibreOffice refuses to open (Excel opens
  it); the content compares correctly either way.
- Not compared, by design: per-row `spans` (an optimisation hint per the spec) and `count`
  attributes (derived from lists that are compared).
