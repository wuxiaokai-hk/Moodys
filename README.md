# scenario_compare — compare two Moody's scenarios quickly

One Python file, **standard library only** (no pip, no openpyxl/pandas). Python 3.8+.

It answers these questions across ~20,000 series:

1. **What are the key changes between scenario A and B?** Every series is ranked by a unit-free *score* (largest gap ÷ the series' typical period-to-period move), so GDP in $bn and a rate in % sit in one ranking.
2. **Where does each variable's gap come from?** For every equation, the gap (B − A) of the variable is split into
   - one line per **input** of its formula: the gap you would get if *only* that input moved from A to B, with the equation re-solved through time. Effects that build up through the variable's own lags are credited to the input that caused them.
   - the **add-factor**: the equation is the same in both scenarios, so whatever the formula does not explain comes from the add-factor (`GDP.US_A`-style series, found automatically with `--af-suffix _A`, or the residual if there is no such series).
3. **What did the scenario actually change?** Series are tagged `X` assumption (no equation), `F` add-factor series, `A` mostly add-factor, `E` formula, `N` not evaluable. The *Where B starts* list and the *Formula vs add-factor* tab show the real shocks.
4. **Can X affect Y, and by how much?** `impact X Y` re-solves every equation between X and Y with **only X changed** (to its scenario-B path, or by a hypothetical `--shock 1%` / `--shock 0.25`). It gives the exact effect on Y and its share of Y's gap. `attribute Y` splits Y's gap across *all* upstream scenario shocks.

Two views of the same gap:
- `explain` = **immediate inputs**. Example: consumption's gap is 25% income, 76% its own add-factor.
- `attribute` = **root shocks**. Example: consumption's gap is 99% the consumption add-factor, because the add-factor also lowers GDP and therefore income.

## Quick start

```bash
# interactive report, one self-contained .html
python scenario_compare.py report --a baseline.csv --b s3.csv --eq model_equations.txt --name-a Baseline --name-b S3 -o s3.html

# same report in your browser, PLUS exact "impact X -> Y" and "split gap across shocks" on demand
python scenario_compare.py serve  --a baseline.csv --b s3.csv --eq model_equations.txt

# terminal
python scenario_compare.py explain FLBR.IDE        --a ... --b ... --eq ...   # inputs + add-factor -> gap
python scenario_compare.py addfactors              --a ... --b ... --eq ...   # formula vs add-factor, all equations
python scenario_compare.py impact FOIL FLBR.IDE    --a ... --b ... --eq ...   # how much X moves Y
python scenario_compare.py impact FRFED FHPI.IOH --shock 1 ...                # +1 unit hypothetical shock
python scenario_compare.py attribute FHPI.IOH      --a ... --b ... --eq ...   # Y's gap split across all shocks
python scenario_compare.py top -n 40 --csv all.csv --a ... --b ... --eq ...   # ranking (+ formula / add-factor shares)
python scenario_compare.py path FOIL FHPI.IOH      --a ... --b ... --eq ...   # the channels
python scenario_compare.py info                    --a ... --b ... --eq ...   # check parsing first
```

Use `py` instead of `python` on Windows, and quote names that contain `$` in bash (`'FGDP$.IUSA'`). `serve` only runs on your own machine (127.0.0.1) and uses nothing outside the standard library.

## Inputs

**Scenario data**: `.csv`, `.tsv` or `.xlsx` (not old `.xls`). The layout is detected automatically:
- series in rows (`Mnemonic, Description, 2000Q1, 2000Q2, …`)
- series in columns with dates down the first column, including Data Buffet-style header rows (`Description:`, `Mnemonic:`, `Source:` …)
- long format (`mnemonic, date, value`)

Date styles accepted: `2026Q1`, `Q1 2026`, `2026:1`, `2026M03`, `2026-03-31`, `3/31/2026`, `Mar-2026`, `2026`, and Excel date serials. Use `--a f1.csv --a f2.csv` to combine several files into one scenario. Use `--sheet` to pick an xlsx sheet and `--meta` to add a CSV of descriptions.

**Equations** (`--eq`): text in EViews syntax: a file, several files, or a folder (`.txt .prg .eq .mod`). The parser reads:
- `Y = …`, `@IDENTITY Y = …`, `series Y = …`, equations that wrap across lines
- `LOG, EXP, ABS, SQR, D(x[,n[,s]]), DLOG, @PCH, @PC, @PCHY, @PCY, @PCA, @MOVAV, @MOVSUM, @TREND`, lags `X(-1)`
- any left-hand side, e.g. `DLOG(Y)` or `LOG(Y/P)`
- spec lines like `eq_y.ls dlog(y) c dlog(x)` (dependencies only)

Equations need **substituted coefficients** (numbers, not `C(1)`) to be decomposed. In EViews, an equation's *Representations* view shows a "Substituted Coefficients" line, and a model's source/text view lists its equations. If your model text only shows links (`:EQ_GDP`), export the substituted form of the equations. Lines with `C(i)` still count for tracing, and `info` tells you how many there are.

## Reading the numbers

- **Window**: from the first period where any series differs to the end of the data (override with `--start/--end`). Contributions are shown at the window end by default (`--at 2027Q4`, or the period selector in the report).
- **Contributions** are in the variable's own units and add up to the gap, apart from a small *interaction* term from non-linearity.
- **Add-factor** = the effect of switching the equation's residual (left-hand side minus formula) from A to B. When a `<name>_A` series exists, it is named and its change is shown. Residual changes that are only rounding noise are ignored.
- **Impact / attribute** hold the add-factors of every other equation at their A values and re-solve the equations period by period. Same-period loops are iterated to convergence. Variables whose equation can't be evaluated are held fixed and listed.
- **Channel share** in `path` = product of the link shares. It is a rough guide to which channel matters; `impact` gives the exact number.

## Options worth knowing

| option | default | what |
|---|---|---|
| `--tol` | 1e-6 | relative tolerance before a series counts as different |
| `--geo-sep` | `.` | character before the geography code (for the *By concept / By geography* tabs) |
| `--af-suffix` | `_A` | suffix of add-factor series (`GDP.US_A`; also matches `GDP_A.US`); comma-separate several |
| `--adj-threshold` | 0.25 | add-factor share that tags a variable "mostly add-factor" |
| `--embed` | 1500 | series with charts and per-period detail in the HTML; the rest are still ranked and traceable |
| `--focus VAR` | | always include VAR in full detail |

## Demo

`make_demo.py` builds a toy Moody's-like model (national block + 50 states, 1,162 series): a baseline CSV, an alternative XLSX (oil +30% fading, policy rate +100bp, add-factors `FC$.IUSA_A` and `FHPI.ICA_A`) and the equations. `demo_report.html` is the output. The tool recovers exactly the four planted shocks as *Where B starts*.
