---
name: jet-reports
description: Use when a Jet Reports workbook (Jet Excel add-in, NL/NF/NP/GL formulas, Excel reports on Microsoft Dynamics NAV / Business Central) has to be created, extended or reviewed, or when an existing Jet report .xlsx must be understood.
---

# Building Jet Reports

## Overview
A Jet report is a normal .xlsx with `_xll.NL(...)` formulas that the Jet add-in evaluates on
**Run Report**. Write it with `scripts/jetgen.py` from a YAML spec – never hand-roll formulas with openpyxl:
Excel 365 stores Jet functions as dynamic-array XLL calls, and `jetgen` reproduces exactly that format
(verified with Jet 23.4, Excel 365, NAV 2017).

## Workflow
1. **Define the numbers first.** Which table and field is the source of truth for each figure (e.g. sales
   from Value Entries vs. Customer Ledger Entries)? If the business has a reference (BI model, finance
   report), match its definition and say where Jet cannot.
2. **Learn the environment's spelling.** Open one existing report of the user (or insert a function with
   Jet's Function Wizard) and copy its style: table/field names with or without numbers, UI language of
   keywords (`Auto+Hide` vs. localized), company key (`Company=` or localized). Put that in a profile
   (see `references/example-profile.yaml`).
3. **Write the spec** following `references/example-spec.yaml`. Company, periods and filters are **options**.
4. **Check:** `python scripts/jetgen.py --check spec.yaml [--profile p.yaml]` must print `OK`.
5. **Build:** `python scripts/jetgen.py spec.yaml JR_Name.xlsx [--profile p.yaml]`; the user runs
   **Jet → Run Report** and reports back. Don't call it done before it ran in Jet.
6. **Tie out** at least one total against the reference report.

## Spec essentials

| Element | Rule |
|---------|------|
| Formulas | Excel English + comma (`IF(a,b,c)`), Jet functions **without** `_xll.` |
| `[[Header]]` | cell of that column in the same row (`$D7`) |
| First table column | the `NL("Rows",…)` replicator |
| Keys | `NL("First",table,,…)` with empty field → record key, then `NF(key,"field")` |
| Exact match | `"@@"&value` whenever a value may contain `*`, `|`, `.`, `..` |
| Company | in **every** NL; one company per NL |
| Mode | `design` (formulas kept) by default; `values` only for distribution |

## Recipes

**Period as option** – two date options plus a computed filter, used as `Period` in NL filters:
```yaml
options:  [{name: From, title: Posting date from, value: 2026-01-01, format: date},
           {name: To,   title: Posting date to,   value: 2026-09-30, format: date}]
computed: [{name: Period, title: Date filter, formula: 'NP("DateFilter",From,To)'}]
```

**Top-N** – sheet `Data` replicates all rows with an amount column and a hidden sort column
(`[[Amount]]-ROW()/1E9` breaks ties); sheet `Top` is a fixed block (`static_rows: 20`, `{n}` = 1…20):
```yaml
- name: Top
  table:
    static_rows: 20
    columns:
      - {header: Rank, formula: '{n}'}
      - {header: Row, hide: true, formula: 'IFERROR(MATCH(LARGE(Data!$H:$H,[[Rank]]),Data!$H:$H,0),"")'}
      - {header: Customer, formula: 'IF([[Row]]="","",INDEX(Data!$C:$C,[[Row]]))'}
```

**Self-check** – put a control cell next to every replicated total: sum of rows minus the total from
one separate `NL("Sum",…)` with the same filters must be 0. Tell the user to check it after the run.

**No access to the data?** Build anyway, list every unverified field/filter value for the user to
confirm in Jet's field list, and leave the tie-out as an explicit open item.

## Common mistakes
- Hard-coding company/period in report cells instead of options.
- Sorting/Top-N on a FlowField with a date FlowFilter – unverified; prefer summing per row and ranking in Excel.
- Localized Excel function names (`WENN`, `SUMME`) – invalid in the file format.
- NAV DateTime fields are UTC; convert before comparing with local times.
- NAV blank date is `01.01.1753`.

See `references/functions.md` for NL/NF/NP syntax and keywords.
