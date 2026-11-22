"""
building_results.py — Export meter-data results to an Excel workbook.

Reads a meter CSV (via door-1-meters/meter.py), then writes an .xlsx with:

  * Sheet "Summary"     — headline stats (readings, span, totals, peaks)
  * Sheet "Top Peaks"   — 3 highest-flow-rate readings with timestamps
  * Sheet "Daily"       — gallons per calendar day
  * Sheet "Flow"        — full time series (timestamp + flow_gpm)
  * Sheet "Chart"       — native Excel line chart of time vs flow rate

Usage (CLI):
    python building_results.py path/to/meter.csv
    python building_results.py meter.csv --unit gph
    python building_results.py meter.csv --out results.xlsx
    python building_results.py meter.csv --building test_building_655

If --building is given, that name is embedded in the summary sheet so the
export is self-describing.

Usage (as a module):
    from building_results import export_building_results
    export_building_results("meter.csv", out="results.xlsx")
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
    from openpyxl.chart import LineChart, Reference
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.stderr.write("openpyxl is required. Install with: pip install openpyxl\n")
    raise


REPO_ROOT = Path(__file__).resolve().parents[1]


def _import_meter_module():
    """Import door-1-meters/meter.py without needing a package."""
    meter_dir = REPO_ROOT / "door-1-meters"
    if str(meter_dir) not in sys.path:
        sys.path.insert(0, str(meter_dir))
    import meter  # type: ignore
    return meter


def top_peaks(df: pd.DataFrame, n: int = 3) -> pd.DataFrame:
    """Top-n rows by flow_gpm with timestamp + flow."""
    return (
        df.sort_values("flow_gpm", ascending=False)
          .head(n)[["timestamp", "flow_gpm"]]
          .reset_index(drop=True)
    )


def _write_kv_sheet(ws, title: str, rows: list[tuple[str, Any]]) -> None:
    ws.title = title
    ws.append(["field", "value"])
    for k, v in rows:
        ws.append([k, v])
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 30


def _write_df_sheet(wb, title: str, df: pd.DataFrame) -> None:
    ws = wb.create_sheet(title)
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append(list(row))
    for i, _ in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(i)].width = 22


def _add_time_vs_flow_chart(wb, flow_sheet: str, n_rows: int) -> None:
    """Add a dedicated Chart sheet with a line chart referencing the Flow sheet."""
    if n_rows < 2:
        return
    chart_ws = wb.create_sheet("Chart")
    data_ws = wb[flow_sheet]

    chart = LineChart()
    chart.title = "Flow rate over time"
    chart.y_axis.title = "Flow rate (gpm)"
    chart.x_axis.title = "Timestamp"
    chart.height = 15
    chart.width = 28
    chart.legend = None

    # Column B = flow_gpm (header in row 1, data rows 2..n_rows+1)
    data_ref = Reference(data_ws, min_col=2, min_row=1, max_row=n_rows + 1, max_col=2)
    chart.add_data(data_ref, titles_from_data=True)

    # Category axis = timestamps in column A
    cats_ref = Reference(data_ws, min_col=1, min_row=2, max_row=n_rows + 1)
    chart.set_categories(cats_ref)

    chart_ws.add_chart(chart, "B2")


def export_building_results(
    meter_csv: str | Path,
    out: str | Path = "building_results.xlsx",
    unit: str | None = None,
    building_name: str | None = None,
) -> Path:
    """
    Produce the building results workbook. Returns the output path.
    """
    meter = _import_meter_module()

    df = meter.load_meter(meter_csv, unit=unit)
    if df.empty:
        raise ValueError("meter CSV produced no rows")

    summary = meter.summary(df)
    daily = meter.daily_totals(df)
    peaks = top_peaks(df, n=3)

    wb = Workbook()

    # Sheet: Summary
    summary_rows: list[tuple[str, Any]] = []
    if building_name:
        summary_rows.append(("building", building_name))
    summary_rows.append(("source_csv", str(Path(meter_csv).resolve())))
    for k, v in summary.items():
        summary_rows.append((k, v))
    _write_kv_sheet(wb.active, "Summary", summary_rows)

    # Sheet: Top Peaks
    peaks_sheet = peaks.copy()
    peaks_sheet.insert(0, "rank", range(1, len(peaks_sheet) + 1))
    _write_df_sheet(wb, "Top Peaks", peaks_sheet)

    # Sheet: Daily
    _write_df_sheet(wb, "Daily", daily)

    # Sheet: Flow (time series) — used by the chart
    flow = df[["timestamp", "flow_gpm"]].copy()
    _write_df_sheet(wb, "Flow", flow)

    # Sheet: Chart — native Excel line chart
    _add_time_vs_flow_chart(wb, "Flow", n_rows=len(flow))

    out_path = Path(out).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Export building meter results to Excel.")
    ap.add_argument("meter_csv", help="Path to meter CSV.")
    ap.add_argument("--unit", default=None, help="Flow unit override: gpm (default), gph, lpm, l_hr.")
    ap.add_argument("--out", default="building_results.xlsx", help="Output .xlsx path.")
    ap.add_argument("--building", default=None, help="Optional building name to embed in the summary.")
    args = ap.parse_args(argv)

    try:
        path = export_building_results(
            args.meter_csv,
            out=args.out,
            unit=args.unit,
            building_name=args.building,
        )
    except (FileNotFoundError, ValueError) as e:
        sys.stderr.write(f"error: {e}\n")
        return 2
    print(f"wrote {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
