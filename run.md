# run.md — how to run the heat pump project

Everything here assumes you're in the repo root (`heat-pump-project/`) in a
terminal (VS Code: `` Ctrl+` ``).

## One-time setup

```powershell
pip install pandas openpyxl ecoengine
```

That's it. No virtualenv required — but if you want one:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install pandas openpyxl ecoengine
```

The sample meter CSV is already in the repo:
`door-1-meters/Example DHW Flow Trend Hourly & Daily(LOG01A (RAW DATA)).csv`.
Commands below use `$METER` as shorthand — set it once per terminal:

```powershell
$METER = "door-1-meters/Example DHW Flow Trend Hourly & Daily(LOG01A (RAW DATA)).csv"
```

---

## 1. Look at the raw meter data (door-1)

Summary stats (readings, span, total gallons, peak hour, etc.):

```powershell
python door-1-meters/meter.py $METER
```

Per-day gallons:

```powershell
python door-1-meters/meter.py $METER --daily
```

24-hour normalized load shape (fraction of daily demand per hour, sums to 1):

```powershell
python door-1-meters/meter.py $METER --profile
```

Options:

- `--unit gpm|gph|lpm|l_hr` — override auto-detected flow unit.
- `--skip-header N` — if auto-detect misses metadata rows at the top.
- `--write-profile shape.csv` — dump the 24-hour shape to a CSV.

---

## 2. Look at the fixtures sheet (door-2)

List every fixture:

```powershell
python door-2-fixtures/fixtures.py door-2-fixtures/general_purpose_hot_water_requirements.xlsx --list
```

Search:

```powershell
python door-2-fixtures/fixtures.py door-2-fixtures/general_purpose_hot_water_requirements.xlsx --find sink
```

Compute total demand for a selection (`name:count`, comma-separated):

```powershell
python door-2-fixtures/fixtures.py door-2-fixtures/general_purpose_hot_water_requirements.xlsx --select "Bar sink:2,Lavatories:4,Triple compartment sink:1"
```

---

## 3. Size the building (sizing-core)

List known buildings:

```powershell
python sizing-core/size.py --list-buildings
```

Size with ecosizer defaults (ignores meter data, uses 25 gpdpp textbook
assumption):

```powershell
python sizing-core/size.py --building test_building_655
```

Size using the real meter data (derives gpdpp from measured daily draw):

```powershell
python sizing-core/size.py --building test_building_655 --meter $METER
```

Also run the 3-day design-day simulation (writes an interactive Plotly
chart to `sim.html`, open in a browser):

```powershell
python sizing-core/size.py --building test_building_655 --meter $METER --simulate --out sim.html
```

Building configs live in `building-list/<name>.xlsx`. Copy `_template.xlsx`
to add a new building; the `name` field there is what you pass to
`--building`.

---

## 4. Export building results to Excel (outputs)

Writes a workbook with:

- Summary — headline stats
- Top Peaks — 3 highest flow-rate readings with timestamps
- Daily — gallons per day
- Flow — full time series
- Chart — native Excel line chart of time vs flow rate

```powershell
python outputs/building_results.py $METER --building test_building_655 --out outputs/building_results.xlsx
```

Open `outputs/building_results.xlsx` in Excel.

---

## 5. Export viable heat pumps (outputs — needs a models CSV)

Not fully wired until you drop a models list in `options/`. When you do:

```powershell
python outputs/heat_pump_results.py --building test_building_655 --models options/models.csv --out outputs/viable.xlsx
```

With meter data folded in:

```powershell
python outputs/heat_pump_results.py --building test_building_655 --models options/models.csv --meter $METER --out outputs/viable.xlsx
```

Filters models whose capacity meets or exceeds the sizing calc's
`min_capacity_kbtuh`. Models CSV/Excel should have one of these columns:
`capacity_kbtuh`, `heating_capacity_kbtuh`, `rated_capacity_kbtuh`, or
`capacity`.

---

## Typical workflow

For a building you already have meter data and a `building-list/<name>.xlsx`
entry for:

```powershell
# 1. sanity-check the meter data
python door-1-meters/meter.py $METER

# 2. see what ecosizer says
python sizing-core/size.py --building test_building_655 --meter $METER

# 3. make the excel to share
python outputs/building_results.py $METER --building test_building_655 --out outputs/building_results.xlsx
```

---

## Folder map

```
building-list/           per-building config (name, occupancy, zip, temps, schematic)
door-1-meters/           meter.py — read cold-water meter CSV
door-2-fixtures/         fixtures.py — read plumbing/equipment hot-water demand sheets
sizing-core/
  size.py                ecosizer wrapper (building + optional meter → sizing result)
  merge-rule/            (next) per-bed design rule aggregated across buildings
check/                   (next) sanity checks on sizing output
options/                 (next) heat pump models list (options.py reads csv/xlsx)
outputs/
  building_results.py    building meter data → .xlsx with chart
  heat_pump_results.py   sizing + models list → .xlsx of viable models
```
