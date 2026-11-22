"""
meter.py — Door 1: cold-water meter reader.

Reads a time-series CSV from a building cold-water meter (timestamps + flow
rate, typically every 5 minutes for a couple of weeks) and produces the
aggregates that downstream sizing (sizing-core) needs:

  * daily total gallons
  * peak hourly demand (gph)
  * normalized 24-hour load shape (fractional demand per hour, sums to 1)
  * basic stats (count, span, avg daily gpd)

No ecosizer usage here — this folder stays a pure meter preprocessor.
sizing-core will consume these outputs later.

Expected CSV columns (case-insensitive; synonyms are mapped):
    timestamp / time / datetime / ts    -> timestamp
    flow / flow_rate / gpm / gph / lpm  -> flow_rate  (+ remembered unit)

Flow unit is auto-detected from the column name. Override with --unit.

Usage (CLI):
    python meter.py path/to/meter.csv
    python meter.py meter.csv --unit gpm
    python meter.py meter.csv --daily          # per-day totals
    python meter.py meter.csv --profile        # 24h normalized load shape
    python meter.py meter.csv --peak-hour      # peak hourly gph
    python meter.py meter.csv --write-profile loadshape.csv

Usage (as a module):
    from meter import load_meter, daily_totals, peak_hour_gph, load_shape_24h
    df = load_meter("meter.csv")
    peak = peak_hour_gph(df)
    shape = load_shape_24h(df)  # pd.Series indexed 0..23, sums to 1.0
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Any

try:
    import pandas as pd
except ImportError:  # pragma: no cover
    sys.stderr.write("pandas is required. Install with: pip install pandas openpyxl\n")
    raise


# ---------- loading ----------

_TIME_SYNONYMS = {"timestamp", "time", "datetime", "ts", "date_time", "date"}
_FLOW_SYNONYMS = {
    "flow", "flow_rate", "flowrate", "rate",
    "gpm", "gph", "lpm", "l_min", "l_hr",
}

# Column name -> implied unit for the flow reading
_UNIT_HINTS = {
    "gpm": "gpm",
    "gph": "gph",
    "lpm": "lpm",
    "l_min": "lpm",
    "l_hr": "l_hr",
}


def _normalize_col(name: str) -> str:
    s = str(name).strip().lower()
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", "_", s).strip("_")
    return s


def _detect_cols(df: pd.DataFrame) -> tuple[str, str, str | None]:
    """Return (time_col, flow_col, unit_hint_or_None)."""
    time_col = None
    flow_col = None
    unit_hint = None

    for col in df.columns:
        if col in _TIME_SYNONYMS and time_col is None:
            time_col = col
        if col in _FLOW_SYNONYMS and flow_col is None:
            flow_col = col
            unit_hint = _UNIT_HINTS.get(col)

    # Fallbacks: first datetime-looking col, first numeric col
    if time_col is None:
        for col in df.columns:
            if "time" in col or "date" in col:
                time_col = col
                break
    if flow_col is None:
        for col in df.columns:
            if col != time_col and pd.api.types.is_numeric_dtype(df[col]):
                flow_col = col
                break

    if time_col is None or flow_col is None:
        raise ValueError(
            f"Could not detect time/flow columns. Got columns: {list(df.columns)}"
        )
    return time_col, flow_col, unit_hint


def load_meter(
    path: str | Path,
    unit: str | None = None,
) -> pd.DataFrame:
    """
    Load a meter CSV and return a DataFrame with columns:
        timestamp (pd.Timestamp), flow_gpm (float), gallons (float)

    `gallons` is the volume for that reading's interval (right-edge delta
    between consecutive timestamps, clipped to the median interval so a
    data gap doesn't inflate one bin).

    Supported units: gpm (default), gph, lpm, l_hr.
    """
    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(f"No such file: {p}")

    df = pd.read_csv(p)
    df.columns = [_normalize_col(c) for c in df.columns]

    time_col, flow_col, unit_hint = _detect_cols(df)
    unit = (unit or unit_hint or "gpm").lower()

    out = pd.DataFrame({
        "timestamp": pd.to_datetime(df[time_col], errors="coerce"),
        "flow_raw": pd.to_numeric(df[flow_col], errors="coerce"),
    }).dropna().sort_values("timestamp").reset_index(drop=True)

    # Convert to gpm
    if unit == "gpm":
        out["flow_gpm"] = out["flow_raw"]
    elif unit == "gph":
        out["flow_gpm"] = out["flow_raw"] / 60.0
    elif unit == "lpm":
        out["flow_gpm"] = out["flow_raw"] / 3.785411784
    elif unit == "l_hr":
        out["flow_gpm"] = out["flow_raw"] / 3.785411784 / 60.0
    else:
        raise ValueError(f"Unsupported unit '{unit}'. Use gpm, gph, lpm, l_hr.")

    # Per-reading gallons (volume since previous reading).
    # Clip excessive gaps to the median interval so one hole doesn't dominate.
    deltas = out["timestamp"].diff().dt.total_seconds() / 60.0  # minutes
    median_dt = deltas.median()
    if pd.isna(median_dt) or median_dt <= 0:
        median_dt = 5.0
    clipped = deltas.clip(upper=median_dt * 2).fillna(median_dt)
    out["interval_min"] = clipped
    out["gallons"] = out["flow_gpm"] * clipped

    return out[["timestamp", "flow_gpm", "interval_min", "gallons"]]


# ---------- aggregates ----------

def daily_totals(df: pd.DataFrame) -> pd.DataFrame:
    """Return per-calendar-day gallons."""
    return (
        df.assign(date=df["timestamp"].dt.date)
          .groupby("date", as_index=False)["gallons"]
          .sum()
          .rename(columns={"gallons": "gallons_per_day"})
    )


def hourly_totals(df: pd.DataFrame) -> pd.DataFrame:
    """Return gallons per calendar hour across the whole dataset."""
    return (
        df.assign(hour_bin=df["timestamp"].dt.floor("h"))
          .groupby("hour_bin", as_index=False)["gallons"]
          .sum()
          .rename(columns={"gallons": "gallons_in_hour"})
    )


def peak_hour_gph(df: pd.DataFrame) -> float:
    """Peak hourly draw (gallons in that hour)."""
    h = hourly_totals(df)
    return float(h["gallons_in_hour"].max()) if not h.empty else 0.0


def load_shape_24h(df: pd.DataFrame) -> pd.Series:
    """
    Normalized 24-hour load shape: fraction of daily demand per hour-of-day.

    Returns a Series indexed 0..23 that sums to 1.0 (if any flow exists).
    Averaged across all days in the dataset.
    """
    by_hour = (
        df.assign(hour=df["timestamp"].dt.hour)
          .groupby("hour")["gallons"]
          .sum()
          .reindex(range(24), fill_value=0.0)
    )
    total = by_hour.sum()
    if total <= 0:
        return by_hour.astype(float)
    return (by_hour / total).astype(float).rename("fraction")


def summary(df: pd.DataFrame) -> dict[str, Any]:
    """Headline numbers for the dataset."""
    if df.empty:
        return {"readings": 0}
    daily = daily_totals(df)
    return {
        "readings":            int(len(df)),
        "first_timestamp":     str(df["timestamp"].min()),
        "last_timestamp":      str(df["timestamp"].max()),
        "span_days":           round((df["timestamp"].max() - df["timestamp"].min()).total_seconds() / 86400.0, 2),
        "median_interval_min": round(float(df["interval_min"].median()), 2),
        "total_gallons":       round(float(df["gallons"].sum()), 1),
        "avg_gpd":             round(float(daily["gallons_per_day"].mean()), 1),
        "max_gpd":             round(float(daily["gallons_per_day"].max()), 1),
        "peak_hour_gph":       round(peak_hour_gph(df), 1),
    }


# ---------- CLI ----------

def _print_dict(d: dict[str, Any]) -> None:
    width = max((len(k) for k in d), default=0)
    for k, v in d.items():
        print(f"{k.ljust(width)}  {v}")


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Door 1 — meter CSV reader.")
    ap.add_argument("path", help="Path to meter CSV.")
    ap.add_argument("--unit", default=None, help="Flow unit: gpm (default), gph, lpm, l_hr.")
    ap.add_argument("--daily", action="store_true", help="Print per-day totals.")
    ap.add_argument("--profile", action="store_true", help="Print 24h normalized load shape.")
    ap.add_argument("--peak-hour", action="store_true", help="Print peak hourly gph only.")
    ap.add_argument("--write-profile", default="", help="Write 24h load shape to this CSV.")
    args = ap.parse_args(argv)

    try:
        df = load_meter(args.path, unit=args.unit)
    except (FileNotFoundError, ValueError) as e:
        sys.stderr.write(f"error: {e}\n")
        return 2

    if args.peak_hour:
        print(round(peak_hour_gph(df), 1))
        return 0

    if args.daily:
        print(daily_totals(df).to_string(index=False))
        return 0

    if args.profile or args.write_profile:
        shape = load_shape_24h(df)
        shape_df = shape.reset_index().rename(columns={"hour": "hour", "fraction": "fraction"})
        if args.write_profile:
            shape_df.to_csv(args.write_profile, index=False)
            print(f"wrote {args.write_profile}")
        if args.profile:
            print(shape_df.to_string(index=False))
        return 0

    _print_dict(summary(df))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
