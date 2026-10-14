# Jet functions – quick reference

> Sources: Jet Reports docs (docs.jet.insightsoftware.com → "Building Functions for Jet Reports",
> "Jet Excel add-in Keywords", "Manually Creating Report Options") and analysis of ~30 production reports
> (≈ 0.5 M NL calls). Items marked *(unverified)* were not tested in Jet.

## NL(What, Table, Field, FilterField1, Filter1, …)

| What | Purpose |
|------|---------|
| `"Rows"` | one Excel row per record (replicator) – `"Rows=3"` = 3 rows per record |
| `"First"` / `"Last"` | first/last record; empty Field → record key for `NF` |
| `"Sum"` | sum of a field |
| `"Count"` / `"CountUnique"` | count / distinct count |
| `"Filter"` | result set as a filter expression for other NL calls |
| `"Lookup"` | pick list in the Report Options dialog |
| `"Sheets"` / `"Columns"` | replicate sheets / columns |

A localized Jet accepts localized What-words too (German: `"Zeilen"`, `"Summe"`, `"Blätter"`, `"Spalten"`).

Special filter fields: `"Company="` (localized e.g. `"Mandant="`), `"Link=",<table>,<field>,"=<field of
main table>",…` (filters after `Link=` apply to the linked table), `"Filters="` (array of pairs, > 10
filters), `"Limit="`, `"DataSource="`, `"Headers="`, `"IncludeDuplicates="`.

Filter values use NAV syntax: `*`, `<>0`, `A|B`, `<>X&<>Y`, `01.01.2026..31.12.2026`;
`"@@"&value` = exact match. Sorting: pass `"+Field","*"` / `"-Field","*"` as a filter pair.

Table/field names: name, or number + name (`"27 Item"`, `"1 No."`) – whatever the Function Wizard
inserts in your environment. Numbers survive caption/language changes.

## NF(Key, Field)
Reads another field of the record returned by an NL with empty Field. One NL key, many NF – cheaper.

## NP(Function, …)
`"DateFilter"` (from, to → filter text), `"Join"` (concatenate), `"Union"` / `"Intersect"` (filter sets),
`"Eval"` (evaluate an Excel formula at run time), company list for a company loop *(name localized,
e.g. `"Mandanten"`)*.

## Keywords

| Where | Keyword | Effect |
|-------|---------|--------|
| A1 of each Jet sheet | `Auto+Hide` | run on Run Report, hide control row/column |
| | `+Values` | convert results to values (formulas backed up) |
| | `+HideSheet` | hide sheet after run (options, version, helpers) |
| | `Title+Fit` | *(unverified details)* |
| Row 1 | `Hide` / `Fit` | hide / autofit column |
| Column A | `Hide` / `Show` | hide row; as formula `=IF(cond,"Hide","Show")` |
| Options sheet row 1 | `Title`, `Value`, `Lookup`, `Tooltip` | Report Options columns |
| Options sheet column A | `Option` | this row is a parameter |

A localized Jet writes localized keywords (German UI: `Auto+Ausblenden`, `+Werte`, `Titel`, `Wert`).

## File format notes
- Excel 365 stores Jet calls as `_xll.NL(...)` dynamic arrays (`cm="1"` + `xl/metadata.xml`).
  Older reports may show `[1]!NL(...)` with an external link to `JetReports.xlam`.
- Reports saved in values mode keep the original formulas in very-hidden sheets.
- The data source is not in the formulas; Jet uses the user's default data source.
