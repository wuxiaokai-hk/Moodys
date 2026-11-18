# scenario_compare — compare two macro scenarios quickly

One Python file, **standard library only** (no pip, no openpyxl/pandas). Python 3.8+.

It answers these questions across ~20,000 series:

1. **What are the key changes between scenario A and B?** Every series is ranked by a unit-free *score* (largest gap ÷ the series' typical period-to-period move), so GDP in $bn and a rate in % sit in one ranking.
2. **Where does each variable's gap come from?** For every equation, the gap (B − A) of the variable is split into
   - one line per **input** of its formula: the gap you would get if *only* that input moved from A to B, with the equation re-solved through time. Effects that build up through the variable's own lags are credited to the input that caused them.
   - the **add-factor**: the equation is the same in both scenarios, so whatever the formula does not explain comes from the add-factor (`GDP.US_A`-style series, found automatically with `--af-suffix _A`, or the residual if there is no such series).
3. **What did the scenario actually change?** Series are tagged `X` assumption (no equation), `F` add-factor series, `A` mostly add-factor, `E` formula, `N` not evaluable. The *Where B starts* list and the *Formula vs add-factor* tab show the real shocks.
4. **Can X affect Y, and by how much?** `impact X Y` re-solves every equation between X and Y with **only X changed** (to its scenario-B path, or by a hypothetical `--shock 1%` / `--shock 0.25`). It gives the exact effect on Y and its share of Y's gap. `attribute Y` splits Y's gap across *all* upstream scenario shocks.

5. **If one variable changes, how many variables are affected, and by how much?** The *Influence* tab and `influence X` re-run the change through the equations and report:
   - how many variables are **linked** to X (use it directly or indirectly)
   - how many actually **move** (≥0.01% of their level), split into ≥1%, 0.1–1% and smaller
   - the most affected variables, with the size and timing of the largest effect

   The report precomputes this for each scenario shock (`--influence-n`, default 20). With `serve` you can run it for any variable and any shock (+1%, +1 unit). Add-factor shocks are applied through the add-factor, and the equation is re-solved.

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
python scenario_compare.py influence               --a ... --b ... --eq ...   # every shock: how many variables move, how much
python scenario_compare.py influence FRFED --shock 1 --a ... --b ... --eq ...  # one variable, hypothetical +1 unit
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

## What's in the report

- **Summary** (first tab): the transmission map can be zoomed (scroll or +/−) and moved (drag). A written summary of how B differs from A, plus:
  - the scenario's shocks, ranked by how much of the key variables they explain;
  - the **key variables**: the changed variables most others depend on, one per concept (or your own with `--key FGDP$.IUSA,FLBR.IUSA,…`);
  - the share of each key variable's gap coming from each shock, with a quarter slider;
  - a **transmission map**: shocks on the left, key variables to the right, in the order the shocks reach them, coloured by their gap. Line thickness = share of the receiving variable's gap; the slider changes the quarter.
- **Impact X → Y**: a quarter slider, a chart of the effect and of its share of Y's gap over time, and the variables in between at the chosen quarter. With `serve` this works for any X and Y; in the static file it works for scenario shocks → key variables.

- **Key changes**: the complete list of series that differ. You can filter it, page through all of it, and download it as a CSV. Each series is tagged *changed directly* (an assumption or add-factor changed) or *via formula only* (same equation, same add-factor; it moved only because its inputs moved).
- **Formula vs add-factor**: the formula/add-factor split for every equation. A time slider picks any quarter, or choose the whole window.
- **By concept / By geography**: count only *directly changed* series as changed. Formula-only movers are shown in a separate column. Click a row to open its full list.
- **Layers**: how many layers of variables each variable affects (longest chain of equations), inputs into an area, and a layer map. See below.
- **Variable detail**: a stacked chart of each input's contribution to the gap in every quarter, a chart of the formula vs add-factor share over time, and a quarter slider for the exact numbers.

## Reading the numbers

- **Window**: from the first period where any series differs to the end of the data (override with `--start/--end`). Contributions are shown at the window end by default (`--at 2027Q4`, or the period selector in the report).
- **Contributions** are in the variable's own units and add up to the gap, apart from a small *interaction* term from non-linearity.
- **Add-factor** = the effect of switching the equation's residual (left-hand side minus formula) from A to B. When a `<name>_A` series exists, it is named and its change is shown. Residual changes that are only rounding noise are ignored.
- **Impact / attribute** hold the add-factors of every other equation at their A values and re-solve the equations period by period. Same-period loops are iterated to convergence. Variables whose equation can't be evaluated are held fixed and listed.
- **Channel share** in `path` = product of the link shares. It is a rough guide to which channel matters; `impact` gives the exact number.

## Which shocks matter most (Importance)

`python scenario_compare.py importance --a … --b … --eq … --csv importance.csv` (also the *Importance* tab of the report).

**Shocks** are the changed exogenous variables (no equation, or listed in `--exog`) and the changed add-factors. Each shock is switched from A to B on its own and everything it feeds is re-solved. For each variable y that moved, its share of y's gap is computed over the whole window:

    share = Σ_t effect_t · gap_t / Σ_t gap_t²

For each variable, the shares of all shocks add up to about 100% (exactly 100% when the model is linear). **Importance** is the average share across variables:
- `--weight concept` (default): each concept counts once, so 50 state copies of one variable don't dominate
- `--weight all`: every variable counts equally
- `--weight key`: only the key variables from the Summary tab (or `--key …`)

The output also gives how many variables a shock explains ≥10% / ≥50% of, and which variables it explains most. The CSV comes with `importance_matrix.csv`, listing every variable × shock share ≥1%.

**Two methods:**
- `--method fast` (default): traces each shock through the model using each equation's own split of its gap into inputs and add-factor (the same split as *Where the gap comes from*). `S(y,s) = Σ_k w_k(y)·S(k,s)`, with the model's loops solved by iteration. One pass over the model whatever the number of shocks (about 6 s on 20,000 equations after loading). It is exact for a linear model; on the demo, importance scores are within 1–2 points of the exact method and the ranking is the same.
- `--method exact`: re-solves the model once per shock. Use it to confirm the top few: `--method exact --max-shocks 20`.

**One geography only:** `importance --geo ICAN` (or the *Variables to explain* selector in the report) averages each shock's share only over that geography's variables. It answers "what drives the Canada forecast?". National or other-country shocks still rank if they move that geography. `--match REGEX` limits to variables by name or description instead.

The report uses the fast method; `--importance-method exact` switches it, and `--no-importance` skips the tab.

**Importance as a report:** add `-o` to `importance` to also write the full HTML report, opened on the Importance tab with your `--geo` and `--weight` already selected (same numbers as the terminal; nothing is computed twice):

```bash
python scenario_compare.py importance --a base.csv --b s4.csv --eq model.txt --geo ICAN -o s4.html
```


## How many layers a variable affects (Layers tab)

**Layers** is the longest chain of equations from a variable downstream. 0 means no equation uses it. 3 means it feeds a variable, which feeds a variable, which feeds a variable. Lags count as links. A loop (GDP → income → consumption → GDP) counts as **one** layer, because inside a loop the longest chain never ends. This describes the model's structure, not the size of an effect.

The tab has:
- a histogram of how many variables reach each number of layers;
- a table with each variable's layers and its full longest chain;
- with an area (`--target` or `--area`), a second number per variable: the longest chain that **ends in the area**, e.g. `--area ICAN`;
- **Inputs into the area**: the variables outside it that its equations use directly, and which area variables each one feeds. Everything else outside reaches the area only through these;
- **Layer maps**, three of them by complexity. **Map 1** is the core: variables with 5+ layers (the cut-offs adjust). **Map 2** is the next tier (3–4 layers) that Map 1 feeds directly. **Map 3** is the rest (1–2 layers) fed by Map 2. Dashed boxes on the left of a map are the variables from the map above that feed it. Click a box to narrow the next maps to what that variable reaches; double-click for its details. Each map shows up to 200 boxes (adjustable), keeping those with the most layers and largest gaps. Scroll to zoom, drag to move.

With `--target`, layers are counted on the full model before variables are dropped, so the overall numbers stay true.

```bash
python scenario_compare.py layers --a base.csv --b s4.csv --eq model.txt --area ICAN            # histogram, top chains, inputs into ICAN
python scenario_compare.py layers --a base.csv --b s4.csv --eq model.txt --area ICAN --by-area  # rank by chains ending in ICAN
```

## Equations that cannot be used

`python scenario_compare.py check --a … --b … --eq … -o not_evaluable.csv` writes one row per equation the tool cannot use, with the reason:
- coefficients not substituted (`C(1)`)
- specification only (`ls y c x`)
- could not parse
- unsupported function (the function is named)
- input series missing from the data (the series is named)
- equation gives missing values (e.g. log of a negative number)
- missing values in some quarters (a warning; the equation is still used)

`report` writes the same file automatically, next to the report, as `<report>_not_evaluable.csv`.

Functions the tool understands: all EViews functions that can appear in a series or model equation (about 170, taken from the EViews Command Reference), including:
- **math:** `LOG, EXP, ABS, SQRT, LOG10, @LOGX, @INV, @FACT, @EXPM1, @LOG1P, @SIGN, @LOGIT, @GAMMA, @MOD, @POW, @ROUND/@FLOOR/@CEILING(x[,n])`
- **trigonometry:** `@SIN, @COS, …`
- **distributions:** `@CNORM, @DNORM, @QNORM, @CLOGISTIC, @ERF, …`
- **logic and NA handling:** `@RECODE, @IIF/@IFF, @NAN, @ISNA, @EQNA, @NEQNA, @BOUNDS, @BETWEEN, @INLIST`, comparisons and `AND/OR`
- **time series:** `D, DLOG, @LAG, @PC/@PCH/@PCA/@PCHA/@PCY/@PCHY, @BRIDGE, @YTD, @DEMEAN, @DETREND`
- **dates:** `@TREND, @TRENDC, @TRENDBR, @OBSNUM, @DURING, @ISPERIOD, @BEFORE, @AFTER, @SEAS, @YEAR, @QUARTER, @MONTH, @ELEM`
- **all moving statistics:** `@MOV*` (NAs propagate) and `@M*` (NAs skipped): sum, average, centred average, stdev, var, cov, cor, max, min, median, …
- **all cumulative statistics:** `@CUM*` and `@CUMB*`, with an optional sample
- **sample statistics:** `@MEAN, @SUM, @STDEV, @MAX, @OBS, @QUANTILE, @FIRST, @LAST, …(x[,"sample"])`

**Sharing a failing equation without its names:** `python scenario_compare.py mask --eq model.txt --failed` prints the equation lines that cannot be used with every series name replaced by `x` and every number by `9`. EViews functions (`@MOVAV` …) and keywords (`AR`, `UNCOND`, `DLOG` …) stay visible, and hidden or unusual characters are spelled out, e.g. `<U+200B zero width space>`. Without `--failed` it shows all lines with `[...]` or AR/MA terms. `--match REGEX` picks other lines.

String, matrix and program functions are not included, because they can't appear in a model equation.

## Other models (not Moody's), and scenarios with different equations

Nothing in the tool is Moody's-specific. Any model whose equations are written as `Y = formula` (EViews style) works, with or without geographies; without geography codes, those tabs simply don't appear.

**Different equations in A and B:** give each scenario its own equation file(s):

```bash
python scenario_compare.py report --a base.csv --b stress.csv \
       --eq-a equations_base.txt --eq-b equations_stress.txt -o compare.html
```

The two sets are compared variable by variable. Spacing, upper/lower case and `[ar(1)=…]` options are ignored:

| Status | Meaning | How its gap is split |
|---|---|---|
| changed | the formula differs | inputs + **equation changed**. Switching to B's equation, with its add-factor, is one effect. Every effect is averaged over A's and B's equation, so an input that only one formula uses still gets its share. |
| only in A | B sets the variable directly | inputs (A's formula) + **equation dropped in B** |
| only in B | A sets the variable directly | inputs (B's formula) + **equation added in B** |

These switches behave like shocks everywhere: Where B starts, `attribute`, Importance and the Summary map. A new tab, **Equations A vs B**, lists every difference with the changed parts highlighted, and the detail panel shows both equations.

**Potential exogenous:** any variable that no equation computes is labelled *potential exogenous* and treated as set directly by the scenario. Equations that cannot be evaluated are treated the same way, and listed in `<report>_not_evaluable.csv`.

Demo: `python make_demo_generic.py demo_generic` builds a bank stress model (no geographies; PD and RWA re-specified, OPEX and DEPOSITS switch between equation and given path); `generic_report.html` is its report.

## Variables the modeller set exogenously

If the modeller overrode a variable that has an equation (exogenized it), list it with `--exog`:

```bash
python scenario_compare.py report --a base.csv --b s3.csv --eq model.txt --exog exog_list.txt
python scenario_compare.py impact FOIL 'FC$.IUSA' ... --exog 'FC$.IUSA,FIR$.IUSA'
```

The list can be a file (names separated by spaces, commas or new lines, `.txt/.csv/.xlsx`) or a comma-separated list.

Listed variables are treated as **assumptions**:
- their whole gap is a scenario choice, so no formula breakdown is shown;
- they are never re-solved, so shocks to their inputs stop there;
- `attribute` counts them as root shocks;
- the report shows their equation as *switched off*.

Without `--exog`, an exogenized variable shows up as *mostly add-factor*. The totals are still right, but the formula split and the `impact` / `influence` results would assume the equation is active.

## Options worth knowing

| option | default | what |
|---|---|---|
| `--target SPEC` | | keep only these variables and everything that feeds them (directly or through any chain of equations, lags included). A geography code (`ICAN`), a wildcard (`"*_ICAN"`), names or a comma list; repeatable. Results for the kept variables are unchanged, just faster |
| `--area SPEC` | `--target` | area for the Layers tab: chains that end in these variables (geography code, wildcard or names) |
| `--tol` | 1e-6 | relative tolerance before a series counts as different |
| `--geo-sep` | auto | character before the geography code: `.`, `_`, or auto-detect. A name ending counts as a geography only if 3+ concepts share it |
| `--exog LIST` | | variables set exogenously by the modeller (file or comma list) |
| `--af-suffix` | `_A` | suffix of add-factor series (`GDP.US_A`; also matches `GDP_A.US`); comma-separate several |
| `--adj-threshold` | 0.25 | add-factor share that tags a variable "mostly add-factor" |
| `--embed` | auto | series with full charts in the HTML: a number, `all`, or auto (all if ≤5,000 changed, else top 1,500). Every other variable still gets its per-quarter shares |
| `--focus VAR` | | always include VAR in full detail. A name, a comma list, or a wildcard like `"*_ICAN"` (all changed Canada series); repeatable |
| `--key LIST` | auto | key variables for the Summary tab (file or comma list) |
| `--summary-n` | 8 | how many key variables to pick automatically (0 turns the Summary off) |
| `-o` (importance) | | also write the full report, opened on the Importance tab |
| `--influence-n` | 20 | scenario shocks simulated for the Influence tab (each takes ~5 s on a 20,000-equation model) |

## Demo

Two demos with the same model and numbers, in the two naming styles:
- `demo/` and `demo_report.html`: names like `FGDP$.IUSA` (default `--geo-sep .`)
- `demo_underscore/` and `demo_report_underscore.html`: names like `FGDP$_IUSA`, run with `--geo-sep _`

Build either with `python make_demo.py 50 800 demo` or `python make_demo.py 50 800 demo_underscore underscore`.

`make_demo.py` builds a toy Moody's-like model (national block + 50 states, 1,162 series): a baseline CSV, an alternative XLSX (oil +30% fading, policy rate +100bp, add-factors `FC$.IUSA_A` and `FHPI.ICA_A`) and the equations. `demo_report.html` is the output. The tool recovers exactly the four planted shocks as *Where B starts*.
