# Spreadsheet Reverse Engineer

"That spreadsheet nobody dares to touch? This maps every formula, finds the traps, and writes the escape plan."

A CPU-only Python CLI for **read-only, descriptive analysis of labelled synthetic workbooks**. It inventories cells, parses formulas, maps dependencies, documents structural business rules, identifies smells, and proposes implementation phases. It never evaluates formulas or writes to the input workbook.

## Quickstart

Requires Python 3.10 or later. From the repository root:

```powershell
python -m pip install -e ".[test]"
python -m sre fixtures --out generated
python -m sre parse generated/pricing.xlsx
python -m sre graph generated/pricing.xlsx
python -m sre rules generated/pricing.xlsx
python -m sre smells generated/pricing.xlsx
python -m sre plan generated/pricing.xlsx
python -m sre report generated/pricing.xlsx --format md --out generated/pricing-report.md
python -m sre report generated/clean.xlsx --format json --out generated/clean-report.json
python -m pytest -q
```

The installed `sre` entry point accepts the same arguments as `python -m sre`. The module form also works when your Python Scripts directory is absent from PATH. All analysis commands accept `--out`; otherwise they print to stdout. `parse`, `graph`, `rules`, `smells`, and `plan` emit JSON with the honesty banner. `report` emits Markdown by default or JSON with `--format json`. Reports include the summary, rule catalogue, smell list, and migration plan. Missing or malformed input returns a nonzero status. An output path resolving to the input, including a hardlink alias, is refused.

`fixtures` creates `clean.xlsx` and `pricing.xlsx`, both fictional and labelled on their README sheet. The pricing fixture has 27 populated cells and 11 formulas. Its planted cycle is exactly:

```text
'Pricing'!C8 -> 'Pricing'!C9 -> 'Pricing'!C8
```

The clean fixture has 8 populated cells, one formula, and no smells under the declared heuristics. This does not mean its calculation has been certified.

## Architecture

```text
Workbook bytes (never modified)
  |
  v
openpyxl inventory: sheets, populated cells, formulas, defined names
  |
  v
Tokenizer + expression grammar validator + finite reference resolver
  |
  v
Dependency graph: precedent -> dependent, formula attached to each edge
  |                         |
  v                         v
Rule catalogue          Structural smell detection
  |                         |
  +------------+------------+
               v
Strongly connected components + topological component ordering
               |
               v
Proposed phases, rule citations, proposed smell remediation, blockers
               |
               v
Deterministic JSON or Markdown report with honesty banner
```

`sre/formula.py` validates and describes formulas; `sre/analysis.py` owns inventory, graph, rules, smells, and planning; `sre/fixtures.py` builds synthetic inputs; `sre/cli.py` provides commands and rendering. openpyxl is the declared runtime dependency. pytest is a development dependency. No GPU, ML framework, network service, or Excel installation is used.

## Formula grammar subset

This is structural analysis, not an Excel calculation engine. Formulas are tokenized with openpyxl and checked by a small expression parser:

```text
formula    := "=" expression
expression := operand
            | ("+" | "-") expression
            | "(" expression ")"
            | function "(" [expression {"," expression}] ")"
            | expression "%"
            | expression binary-op expression
operand    := number | quoted-text | logical | error | reference | defined-name
binary-op  := "=" | "<>" | "<" | ">" | "<=" | ">=" | "&"
            | "+" | "-" | "*" | "/" | "^"
reference  := [sheet "!"] A1-cell [":" A1-cell]
```

References support `$` anchors, quoted sheet names, doubled apostrophes, case-insensitive sheet matching, finite rectangles, and workbook or sheet scoped names pointing to supported references. Local names take precedence within their sheet. Names are recursively resolved with recursion checks. Blank referenced cells become graph nodes. Reference endpoints must remain within Excel's grid. Expansion is capped at **100,000 cells per reference**; larger ranges are explicitly flagged. These limits ensure bounded CPU and memory usage.

Supported function names:

```text
SUM AVERAGE MIN MAX COUNT COUNTA
IF IFS AND OR NOT IFERROR
VLOOKUP HLOOKUP XLOOKUP
ROUND ROUNDUP ROUNDDOWN ABS INT MOD TRUE FALSE
NOW TODAY RAND RANDBETWEEN OFFSET INDIRECT CELL INFO
```

Function argument expressions are recorded, including nested calls. This parser validates syntax structure, not function arity, type compatibility, match semantics, calculated values, or financial meaning. `INDIRECT` and `OFFSET` are recognized but always flagged: their dynamic dependencies cannot be statically resolved. Supported literals inside formulas are documented as hardcoded constants even when conventional, such as lookup column indexes or `FALSE` match modes.

Whole-column or whole-row ranges, unions, intersections, structured references, external workbook references, 3D ranges, spilled references, array constants, omitted arguments, unsupported functions, array/data-table formulas, and malformed expressions require manual review. Cell contents remain inventoried where possible. Tables, data validation, conditional formatting, pivot calculations, and external workbook links are reported as outside the cell-formula analysis. Formula dependencies that can be resolved are retained even when other portions are unsupported; **an incomplete graph is explicitly marked** and must not be used as a complete account of the workbook. Chart behavior, macros, and Excel's runtime recalculation semantics are not implemented. The CLI targets `.xlsx` synthetic workbooks.

## Rule catalogue

Every rule has a stable ID and cell citations. The catalogue records:

| Rule | Description |
| --- | --- |
| Formula | Every formula cell and its original expression |
| Named range | Definition, member cells, and the direct dependent cells it feeds |
| Lookup | VLOOKUP first-column search, HLOOKUP first-row search, or XLOOKUP paired lookup/return arrays, with argument expressions and referenced cells |
| Conditional | IF, IFS, and IFERROR branch argument expressions with the formula cell |
| Aggregation | SUM, AVERAGE, MIN, MAX, COUNT, or COUNTA with direct precedents |

These descriptions expose actual expressions without inventing business intent. For example, the fictional `TaxRate` name points to `'Tax'!A2` and feeds `'Pricing'!H2` and `'Pricing'!K2`. The VLOOKUP rule cites the formula, search cell, and all four cells of the fictional Rates table. Chained aggregations remain visible through the graph and their individual aggregation rules.

## Smell catalogue

| Smell | Detection and consequence |
| --- | --- |
| Circular reference | Strongly connected components with more than one node, or a self-edge. Each component cites all affected cells and reports one exact closed cycle witness. Feedback must be explicitly resolved. Enumerating every possible simple cycle is intentionally avoided. |
| Dead cells | Populated numeric or formula cells with no direct dependents, except README metadata and explicitly named outputs. These may be unused logic or intended outputs. Text labels are excluded to avoid flagging headings. |
| Duplicated formula pattern | Equal token structures after converting cell references to relative row/column offsets while preserving absolute anchors, literals, and sheet qualifiers. Candidates for a shared implementation rule. |
| Volatile functions | NOW, TODAY, RAND, RANDBETWEEN, OFFSET, INDIRECT, CELL, INFO. These can recalculate or vary independently of explicit edges. CELL and INFO are conservatively listed regardless of argument-dependent volatility. |
| Hardcoded constants | Number, text, or logical literals embedded in formulas. They may hide assumptions or parameter choices. |
| Cross-sheet spaghetti | A formula directly referencing **at least 3 distinct other sheets**. This declared threshold detects coupling between model areas. |
| Manual review | Unsupported syntax, unresolved references, non-cell logic, or missing synthetic labels. Findings explain the limitation and cite locations where available. |

Names beginning with `Output` (case-insensitive) explicitly declare terminal output cells and exempt them from the dead-cell heuristic. Use this convention to distinguish unused calculations from expected deliverables. Heuristics are descriptive findings, not correctness judgments. A copy-down formula is allowed but still reported as a consolidation opportunity.

## Migration planner

Phases contain all graph nodes, including inputs, blank precedents, metadata, and formulas. Edges always point from an input/precedent to its dependent. All edges between different strongly connected components go from an earlier phase to a later phase. A cyclic component stays together and is blocked; it is never presented as a valid internal topological ordering. Manual-review cells and downstream dependents are also blocked. Workbook-level limitations require review independently of phases.

Each phase lists cited rules and `smells_to_address`, with a remediation action. Rules and smells spanning multiple phases may appear in each applicable phase. All rules are covered. Workbook-level findings without cells appear in `unassigned_smells`. Remediation is proposed and is never claimed to have happened. A plan does not include an application implementation or guarantee equivalent results; reviewed specifications and independent validation would be subsequent work.

## Determinism and testing

No random generator is used, so there is no seed to configure. Fixture document metadata and ZIP entry timestamps are fixed to **2020-01-01 00:00:00**; entries are sorted and compression is fixed. Regeneration with the same dependency versions yields byte-identical workbook files. Results sort IDs, cells, graph edges, names, and component processing. Reports include neither analysis timestamps nor absolute input paths. The basename remains part of the summary.

Tests cover hand-computed dependency edges and exact cycle paths, planted smells, clean-model behavior, named ranges and lookup/conditional citations, dependency-respecting phases and full rule coverage, unsupported constructs, byte-identical input files after analysis, regenerated fixture bytes, report repeatability across Python hash seeds, long dependency chains, all CLI commands, end-to-end Markdown/JSON output, output alias protection, and honesty banners.

## HONESTY BOUNDARY

**Analysis is read-only and descriptive. It documents what formulas structurally do; it does not certify the spreadsheet is correct. The migration plan is a proposed ordering, not a guarantee. All provided workbooks are fictional synthetic fixtures and labelled. This project does not audit real finances.**

Every report carries `SYNTHETIC FIXTURE ANALYSIS ONLY` and the explicit no-audit, no-certification, proposed-ordering banner. The inventory independently records `synthetic_labelled`. An unlabelled input is identified as such and produces a manual-review finding stating that the tool does not audit real finances. The banner never establishes that an arbitrary user file is synthetic. Do not interpret the report as financial advice, certification, or evidence that business logic is sound.

## Local build record

The project is committed locally on `main`. Publishing is a separate verification step and is not performed by this build. `BUILD-DONE.txt` records the passing pytest count and the tested implementation commit immediately before the final build-record commit, avoiding a self-referential commit hash.
