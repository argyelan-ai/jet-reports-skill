# jet-reports – a Claude Code skill for Jet Reports

Lets Claude Code build **Jet Reports** workbooks (the Jet Excel add-in for Microsoft Dynamics NAV /
Business Central): you describe the report, Claude writes a YAML spec, `jetgen.py` checks it and
produces a ready-to-run `.xlsx` with `NL` / `NF` / `NP` formulas, a Report Options sheet and a version log.

Verified end-to-end with **Jet Reports 23.4.0.3181**, Excel 365 (de-CH), Dynamics NAV 2017:
the generated file opens without repair prompts and runs with **Jet → Run Report**.

## Install

```bash
# personal skill (all projects)
cp -r skills/jet-reports ~/.claude/skills/
pip install openpyxl pyyaml
```

## Use

Ask Claude, e.g. *"Build a Jet report: top customers by sales for a selectable period"*.
Or run the generator yourself:

```bash
python ~/.claude/skills/jet-reports/scripts/jetgen.py --check my-spec.yaml
python ~/.claude/skills/jet-reports/scripts/jetgen.py my-spec.yaml JR_MyReport.xlsx --profile my-profile.yaml
```

- `references/example-spec.yaml` – spec format
- `references/example-profile.yaml` – localized Jet (keywords, company key, table numbers)
- `references/functions.md` – NL/NF/NP and keyword cheat sheet

## How it works

Excel 365 stores Jet functions as XLL dynamic-array formulas (`_xll.NL(...)`, `cm="1"`,
`xl/metadata.xml`). openpyxl cannot write that directly, so `jetgen` writes array formulas and
post-processes the package. No Jet binaries are inspected or modified – the tool only writes
standard Excel files that the licensed add-in evaluates.

## Tests

```bash
pip install pytest && pytest tests
```

## Disclaimer

Not affiliated with or endorsed by insightsoftware. "Jet Reports" is a trademark of its owner.
You need a licensed Jet installation to run the generated reports.

## License

MIT – see `LICENSE`.
