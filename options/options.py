"""
options.py — Heat pump options loader.

Reads a list of heat pump models (and their specs) from a local .csv or
Excel file and gives you simple ways to list, filter, and look them up.

Usage (CLI):
    python options.py path/to/file.csv
    python options.py path/to/file.xlsx --sheet "Sheet1"
    python options.py path/to/file.csv --list
    python options.py path/to/file.csv --filter "capacity_btu>=36000,brand=Mitsubishi"
    python options.py path/to/file.csv --model "MXZ-SM48NAMHZ"
    python options.py path/to/file.csv --columns

Usage (as a module):
    from options import load_options, filter_options
    df = load_options("heat_pumps.xlsx")
    hits = filter_options(df, {"capacity_btu": (">=", 36000), "brand": Mitsubishi"})

Notes:
  * Column names are normalized to snake_case so "Capacity (BTU)" becomes
    "capacity_btu". Use --columns to see what you've got after loading.
  * No fixed schema is enforced — bring whatever columns you want. Common
    ones: brand, model, capacity_btu, cop, seer, hspf, price_usd, type.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Any, Iterable

try:
    import pandas as pd
except ImportError:  # pragma: no cover
    sys.stderr.write(
        "pandas is required. Install with: pip install pandas openpyxl\n"
    )
    raise


# ---------- loading ----------

def _normalize_col(name: str) -> str:
    """Convert 'Capacity (BTU/hr)' -> 'capacity_btu_hr'."""
    s = str(name).strip().lower()
    s = re.sub(r"[^\w\s]", " ", s)       # drop punctuation
    s = re.sub(r"\s+", "_", s).strip("_")
    return s


def load_options(path: str | Path, sheet: str | int | None = 0) -> pd.DataFrame:
    """
    Load heat pump options from a CSV or Excel file.

    Parameters
    ----------
    path : str or Path
        Local path to a .csv, .xlsx, .xls, or .xlsm file.
    sheet : str or int or None
        Sheet name/index for Excel files. Ignored for CSV. Default: first sheet.
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
        raise ValueError(
            f"Unsupported file type '{suffix}'. Use .csv, .xlsx, .xls, or .xlsm."
        )

    df.columns = [_normalize_col(c) for c in df.columns]
    return df


# ---------- querying ----------

_OPS = {
    ">=": lambda a, b: a >= b,
    "<=": lambda a, b: a <= b,
    ">":  lambda a, b: a > b,
    "<":  lambda a, b: a < b,
    "!=": lambda a, b: a != b,
    "=":  lambda a, b: a == b,
    "==": lambda a, b: a == b,
}


def _coerce(val: str) -> Any:
    """Turn a string filter value into int/float if possible."""
    for cast in (int, float):
        try:
            return cast(val)
        except ValueError:
            continue
    return val


def parse_filter_string(expr: str) -> dict[str, tuple[str, Any]]:
    """
    Parse a filter expression like:
        "capacity_btu>=36000,brand=Mitsubishi,cop>3.5"
    into:
        {"capacity_btu": (">=", 36000), "brand": ("=", "Mitsubishi"), ...}
    """
    filters: dict[str, tuple[str, Any]] = {}
    if not expr:
        return filters

    pattern = re.compile(r"\s*([\w.]+)\s*(>=|<=|!=|==|=|>|<)\s*([^,]+?)\s*(?:,|$)")
    for col, op, raw in pattern.findall(expr):
        filters[col] = (op, _coerce(raw.strip()))
    return filters


def filter_options(
    df: pd.DataFrame,
    filters: dict[str, tuple[str, Any]] | Iterable[tuple[str, str, Any]],
) -> pd.DataFrame:
    """
    Apply filters to a DataFrame.

    Accepts either:
      * dict:   {"capacity_btu": (">=", 36000), "brand": ("=", "Mitsubishi")}
      * tuples: [("capacity_btu", ">=", 36000), ("brand", "=", "Mitsubishi")]
    """
    if isinstance(filters, dict):
        items = [(c, op, v) for c, (op, v) in filters.items()]
    else:
        items = list(filters)

    result = df
    for col, op, val in items:
        if col not in result.columns:
            raise KeyError(f"Column '{col}' not found. Available: {list(result.columns)}")
        fn = _OPS.get(op)
        if fn is None:
            raise ValueError(f"Unsupported operator '{op}'.")
        if isinstance(val, str) and result[col].dtype == object:
            # case-insensitive string compare
            series = result[col].astype(str).str.casefold()
            target = val.casefold()
            mask = fn(series, target)
        else:
            mask = fn(result[col], val)
        result = result[mask]
    return result


def find_model(df: pd.DataFrame, model_name: str, column: str = "model") -> pd.DataFrame:
    """Return rows whose `column` contains `model_name` (case-insensitive)."""
    if column not in df.columns:
        raise KeyError(f"Column '{column}' not found. Available: {list(df.columns)}")
    mask = df[column].astype(str).str.contains(model_name, case=False, na=False)
    return df[mask]


# ---------- CLI ----------

def _print_df(df: pd.DataFrame, limit: int | None = None) -> None:
    if df.empty:
        print("(no rows)")
        return
    view = df.head(limit) if limit else df
    with pd.option_context("display.max_rows", None, "display.max_columns", None, "display.width", 200):
        print(view.to_string(index=False))


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Heat pump options loader.")
    ap.add_argument("path", help="Path to .csv or Excel file.")
    ap.add_argument("--sheet", default=0, help="Sheet name/index for Excel files.")
    ap.add_argument("--list", action="store_true", help="List all rows.")
    ap.add_argument("--columns", action="store_true", help="Show available columns.")
    ap.add_argument("--filter", default="", help='Filter expr, e.g. "capacity_btu>=36000,brand=Mitsubishi".')
    ap.add_argument("--model", default="", help="Find rows matching a model name (substring, case-insensitive).")
    ap.add_argument("--limit", type=int, default=None, help="Limit number of rows printed.")
    args = ap.parse_args(argv)

    try:
        sheet: Any = args.sheet
        try:
            sheet = int(sheet)
        except (TypeError, ValueError):
            pass
        df = load_options(args.path, sheet=sheet)
    except (FileNotFoundError, ValueError) as e:
        sys.stderr.write(f"error: {e}\n")
        return 2

    if args.columns:
        print("columns:", list(df.columns))
        return 0

    if args.model:
        df = find_model(df, args.model)

    if args.filter:
        df = filter_options(df, parse_filter_string(args.filter))

    if args.list or args.filter or args.model or not (args.columns):
        _print_df(df, args.limit)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
