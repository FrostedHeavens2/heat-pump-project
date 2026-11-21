"""
fixtures.py — Door 2: Plumbing fixtures (hot water demand).

Reads a list of fixtures / equipment and their hot-water demand from a
CSV or Excel sheet. Starts with a small list of kitchen equipment; will
grow to cover all plumbing fixtures.

Expected columns (names are normalized to snake_case; synonyms are mapped):
    equipment / fixture / name  -> equipment
    gph / gal_per_hr            -> gph
    l_hr / lph / lpm            -> l_hr  (lpm is converted to l_hr)

Usage (CLI):
    python fixtures.py path/to/file.xlsx
    python fixtures.py path/to/file.csv --list
    python fixtures.py path/to/file.xlsx --find "sink"
    python fixtures.py path/to/file.xlsx --select "Bar sink:2,Lavatories:4"
    python fixtures.py path/to/file.xlsx --select-file selections.csv
    python fixtures.py path/to/file.xlsx --columns

Usage (as a module):
    from fixtures import load_fixtures, find_fixture, total_demand
    df = load_fixtures("fixtures.xlsx")
    row = find_fixture(df, "bar sink")
    totals = total_demand(df, {"Bar sink": 2, "Lavatories": 4})

A selection file is a tiny 2-column CSV like:
    equipment,count
    Bar sink,2
    Lavatories,4
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Any, Mapping

try:
    import pandas as pd
except ImportError:  # pragma: no cover
    sys.stderr.write("pandas is required. Install with: pip install pandas openpyxl\n")
    raise


# ---------- loading ----------

_COLUMN_SYNONYMS = {
    "equipment": "equipment",
    "fixture": "equipment",
    "fixtures": "equipment",
    "name": "equipment",
    "item": "equipment",
    "description": "equipment",
    "gph": "gph",
    "gal_hr": "gph",
    "gal_per_hr": "gph",
    "gallons_per_hour": "gph",
    "gallons_hr": "gph",
    "l_hr": "l_hr",
    "l_per_hr": "l_hr",
    "liters_per_hour": "l_hr",
    "liters_hr": "l_hr",
    "lph": "l_hr",
}


def _clean_cell(v: Any) -> Any:
    """Strip stray non-printing chars (degree-ish mojibake) and whitespace."""
    if not isinstance(v, str):
        return v
    # Drop replacement chars, NBSPs, and other control / non-ASCII cruft
    cleaned = re.sub(r"[\ufffd\xa0\u00b0]", "", v)
    cleaned = re.sub(r"[^\x20-\x7e]", "", cleaned)
    return cleaned.strip()


def _normalize_col(name: str) -> str:
    s = _clean_cell(str(name)).lower()
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", "_", s).strip("_")
    return _COLUMN_SYNONYMS.get(s, s)


def load_fixtures(path: str | Path, sheet: str | int | None = 0) -> pd.DataFrame:
    """
    Load fixtures from a CSV or Excel file.

    Returns a DataFrame with cleaned text, snake_case columns, and numeric
    gph / l_hr where present.
    """
    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(f"No such file: {p}")

    suffix = p.suffix.lower()
    if suffix == ".csv":
        df = pd.read_csv(p)
    elif suffix in {".xlsx", ".xls", ".xlsm"}:
        df = pd.read_excel(p, sheet_name=sheet)
    else:
        raise ValueError(f"Unsupported file type '{suffix}'. Use .csv or .xlsx.")

    df.columns = [_normalize_col(c) for c in df.columns]

    # Clean string columns
    for col in df.select_dtypes(include=["object", "string"]).columns:
        df[col] = df[col].map(_clean_cell)

    # Coerce numeric where it makes sense
    for col in ("gph", "l_hr"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


# ---------- querying ----------

def find_fixture(df: pd.DataFrame, query: str, column: str = "equipment") -> pd.DataFrame:
    """Case-insensitive substring match on an equipment/name column."""
    if column not in df.columns:
        raise KeyError(f"Column '{column}' not found. Available: {list(df.columns)}")
    mask = df[column].astype(str).str.contains(query, case=False, na=False)
    return df[mask]


def total_demand(
    df: pd.DataFrame,
    selections: Mapping[str, float],
    column: str = "equipment",
) -> dict[str, float]:
    """
    Given a mapping of equipment-name -> count, return total hot-water demand.

    Match is case-insensitive substring. Unknown names are reported.
    Returns a dict like {"gph": 150.0, "l_hr": 567.75, "missing": ["foo"]}.
    """
    totals: dict[str, float] = {"gph": 0.0, "l_hr": 0.0}
    missing: list[str] = []

    for name, count in selections.items():
        hits = find_fixture(df, name, column=column)
        if hits.empty:
            missing.append(name)
            continue
        if len(hits) > 1:
            # Prefer exact (case-insensitive) match if present
            exact = hits[hits[column].astype(str).str.casefold() == name.casefold()]
            hits = exact if not exact.empty else hits.head(1)
        row = hits.iloc[0]
        for unit in ("gph", "l_hr"):
            if unit in df.columns and pd.notna(row.get(unit)):
                totals[unit] += float(row[unit]) * float(count)

    out: dict[str, Any] = {k: round(v, 2) for k, v in totals.items()}
    if missing:
        out["missing"] = missing
    return out


# ---------- selection parsing ----------

def parse_selection_string(expr: str) -> dict[str, float]:
    """Parse 'Bar sink:2, Lavatories:4' into {'Bar sink': 2, 'Lavatories': 4}."""
    out: dict[str, float] = {}
    for part in expr.split(","):
        part = part.strip()
        if not part:
            continue
        if ":" not in part:
            raise ValueError(f"Bad selection '{part}' (expected name:count)")
        name, count = part.rsplit(":", 1)
        out[name.strip()] = float(count.strip())
    return out


def load_selection_file(path: str | Path) -> dict[str, float]:
    """Load a tiny 2-column (equipment,count) CSV of selections."""
    df = pd.read_csv(path)
    df.columns = [_normalize_col(c) for c in df.columns]
    if "equipment" not in df.columns or "count" not in df.columns:
        raise ValueError("selection file must have 'equipment' and 'count' columns")
    return {str(r.equipment).strip(): float(r.count) for r in df.itertuples(index=False)}


# ---------- CLI ----------

def _print_df(df: pd.DataFrame) -> None:
    if df.empty:
        print("(no rows)")
        return
    with pd.option_context("display.max_rows", None, "display.max_columns", None, "display.width", 200):
        print(df.to_string(index=False))


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Door 2 — plumbing fixtures loader.")
    ap.add_argument("path", help="Path to .csv or Excel file.")
    ap.add_argument("--sheet", default=0, help="Sheet name/index for Excel files.")
    ap.add_argument("--list", action="store_true", help="List all fixtures.")
    ap.add_argument("--columns", action="store_true", help="Show available columns.")
    ap.add_argument("--find", default="", help="Substring search in equipment column.")
    ap.add_argument("--select", default="", help='Compute totals, e.g. "Bar sink:2,Lavatories:4".')
    ap.add_argument("--select-file", default="", help="CSV of equipment,count pairs.")
    args = ap.parse_args(argv)

    try:
        sheet: Any = args.sheet
        try:
            sheet = int(sheet)
        except (TypeError, ValueError):
            pass
        df = load_fixtures(args.path, sheet=sheet)
    except (FileNotFoundError, ValueError) as e:
        sys.stderr.write(f"error: {e}\n")
        return 2

    if args.columns:
        print("columns:", list(df.columns))
        return 0

    if args.find:
        _print_df(find_fixture(df, args.find))
        return 0

    selections: dict[str, float] = {}
    if args.select:
        selections.update(parse_selection_string(args.select))
    if args.select_file:
        selections.update(load_selection_file(args.select_file))

    if selections:
        totals = total_demand(df, selections)
        for k, v in totals.items():
            print(f"{k}: {v}")
        return 0

    _print_df(df)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
