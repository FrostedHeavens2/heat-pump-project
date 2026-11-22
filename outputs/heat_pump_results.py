"""
heat_pump_results.py — Export viable heat pump models based on sizing results.

Pipeline:
  1. Run sizing-core/size.py for a building (optionally with meter data) to
     get ecosizer's min_capacity_kbtuh and min_storage_storageT_gal.
  2. Load the heat pump models list from options/options.py.
  3. Filter models whose capacity meets or exceeds the required capacity.
  4. Export the viable list + the sizing results to an Excel workbook.

Usage (CLI):
    python heat_pump_results.py --building test_building_655 --models path/to/models.csv
    python heat_pump_results.py --building test_building_655 --models models.xlsx \
        --meter ../door-1-meters/data.csv --out viable.xlsx

The models file (csv or xlsx) is read with options/options.py, so column
names are normalized to snake_case. For filtering we look for a capacity
column; recognized names (first match wins):
    capacity_kbtuh, capacity_kbtu_hr, heating_capacity_kbtuh,
    rated_capacity_kbtuh, capacity

If none are present, the script lists what columns it saw and exits with
a clear message — adjust the models file or extend CAPACITY_COLUMN_CANDIDATES.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

try:
    import pandas as pd
except ImportError:  # pragma: no cover
    sys.stderr.write("pandas is required. Install with: pip install pandas openpyxl\n")
    raise

try:
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.stderr.write("openpyxl is required. Install with: pip install openpyxl\n")
    raise


REPO_ROOT = Path(__file__).resolve().parents[1]

CAPACITY_COLUMN_CANDIDATES = [
    "capacity_kbtuh",
    "capacity_kbtu_hr",
    "heating_capacity_kbtuh",
    "rated_capacity_kbtuh",
    "capacity",
]


def _import_sibling(folder: str, module: str):
    """Import a module from one of the sibling folders without packaging."""
    target = REPO_ROOT / folder
    if str(target) not in sys.path:
        sys.path.insert(0, str(target))
    return __import__(module)


def _find_capacity_col(df: pd.DataFrame) -> str | None:
    for cand in CAPACITY_COLUMN_CANDIDATES:
        if cand in df.columns:
            return cand
    return None


def filter_viable(models: pd.DataFrame, min_capacity_kbtuh: float) -> pd.DataFrame:
    """Keep models whose capacity meets or exceeds the required capacity."""
    col = _find_capacity_col(models)
    if col is None:
        raise ValueError(
            f"No capacity column found in models. Available: {list(models.columns)}.\n"
            f"Expected one of: {CAPACITY_COLUMN_CANDIDATES}"
        )
    capacity = pd.to_numeric(models[col], errors="coerce")
    return models[capacity >= float(min_capacity_kbtuh)].copy()


def _write_kv_sheet(ws, title: str, rows: list[tuple[str, Any]]) -> None:
    ws.title = title
    ws.append(["field", "value"])
    for k, v in rows:
        ws.append([k, v])
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 36


def _write_df_sheet(wb, title: str, df: pd.DataFrame) -> None:
    ws = wb.create_sheet(title)
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append(list(row))
    for i, _ in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(i)].width = 22


def export_heat_pump_results(
    building: str,
    models_path: str | Path,
    out: str | Path = "heat_pump_results.xlsx",
    meter_csv: str | Path | None = None,
) -> Path:
    """
    Run sizing for `building`, filter the heat pump models that fit, and
    export the result. Returns the output path.
    """
    size_mod = _import_sibling("sizing-core", "size")
    options_mod = _import_sibling("options", "options")

    cfg = size_mod.load_building(building)
    sizing_result = size_mod.run_ecosizer(cfg, meter_csv=meter_csv, simulate=False)
    sizing = sizing_result["sizing"]
    min_capacity = float(sizing["min_capacity_kbtuh"])

    models = options_mod.load_options(models_path)
    viable = filter_viable(models, min_capacity)

    wb = Workbook()
    summary_rows: list[tuple[str, Any]] = [
        ("building",                 sizing_result["building"]),
        ("models_source",            str(Path(models_path).resolve())),
        ("models_total",             int(len(models))),
        ("models_viable",            int(len(viable))),
        ("min_capacity_kbtuh",       round(min_capacity, 2)),
        ("min_storage_storageT_gal", round(float(sizing.get("min_storage_storageT_gal", 0)), 2)),
        ("min_tm_volume_gal",        round(float(sizing.get("min_tm_volume_gal", 0)), 2)),
        ("min_tm_capacity_kbtuh",    round(float(sizing.get("min_tm_capacity_kbtuh", 0)), 2)),
    ]
    if "gpdpp_from_meter" in sizing_result:
        summary_rows.append(("gpdpp_from_meter", sizing_result["gpdpp_from_meter"]))
    _write_kv_sheet(wb.active, "Summary", summary_rows)

    _write_df_sheet(wb, "Viable Models", viable)
    _write_df_sheet(wb, "All Models",    models)

    out_path = Path(out).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Export viable heat pumps for a building.")
    ap.add_argument("--building", required=True, help="Building name (from building-list/) or xlsx path.")
    ap.add_argument("--models",   required=True, help="Path to heat pump models csv/xlsx.")
    ap.add_argument("--meter",    default=None, help="Optional meter CSV for gpdpp derivation.")
    ap.add_argument("--out",      default="heat_pump_results.xlsx", help="Output xlsx path.")
    args = ap.parse_args(argv)

    try:
        path = export_heat_pump_results(
            building=args.building,
            models_path=args.models,
            out=args.out,
            meter_csv=args.meter,
        )
    except (FileNotFoundError, ValueError, RuntimeError) as e:
        sys.stderr.write(f"error: {e}\n")
        return 2
    print(f"wrote {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
