"""
size.py — Sizing-core: ecosizer wrapper.

Takes a building (from ../building-list/<name>.xlsx) and runs Ecotope's
ecosizer engine against it to produce sizing results (capacity, storage,
etc.) and optionally a 3-day design-day simulation.

Big picture: sizing-core is our python engine that turns per-building meter
data into a per-bed design rule. This wrapper is one tool inside that
engine — it drives ecosizer. Later, merge-rule aggregates across buildings
to produce the per-bed design rule, and check validates results.

Usage (CLI):
    python size.py --building test_building_655
    python size.py --building test_building_655 --meter ../door-1-meters/data.csv
    python size.py --building test_building_655 --simulate --out sim.html
    python size.py --list-buildings

The meter CSV (optional, for later) feeds measured daily demand into the
sizing calc via gpdpp = total_gallons_per_day / occupancy.

Usage (as a module):
    from size import load_building, run_ecosizer
    cfg = load_building("test_building_655")
    result = run_ecosizer(cfg, meter_csv=None, simulate=False)
    print(result["sizing"])
"""

from __future__ import annotations

import argparse
import math
import sys
from pathlib import Path
from typing import Any

try:
    import pandas as pd
except ImportError:  # pragma: no cover
    sys.stderr.write("pandas is required. Install with: pip install pandas openpyxl\n")
    raise

try:
    from ecoengine import EcosizerEngine  # type: ignore
    _ECOENGINE_OK = True
    _ECOENGINE_ERR: str | None = None
except Exception as e:  # pragma: no cover
    _ECOENGINE_OK = False
    _ECOENGINE_ERR = str(e)


REPO_ROOT = Path(__file__).resolve().parents[1]
BUILDING_DIR = REPO_ROOT / "building-list"


# ---------- building loader ----------

def load_building(name_or_path: str) -> dict[str, Any]:
    """
    Load a building config from building-list/.

    Accepts either a bare name ("test_building_655") or a full path.
    Returns a dict of {field: value} with numeric fields coerced.
    """
    p = Path(name_or_path)
    if not p.exists():
        candidate = BUILDING_DIR / f"{name_or_path}.xlsx"
        if candidate.exists():
            p = candidate
        else:
            raise FileNotFoundError(
                f"Building '{name_or_path}' not found in {BUILDING_DIR} "
                f"or as a direct path."
            )

    df = pd.read_excel(p)
    df.columns = [str(c).strip().lower() for c in df.columns]
    if "field" not in df.columns or "value" not in df.columns:
        raise ValueError(
            f"{p.name}: expected columns 'field' and 'value', got {list(df.columns)}"
        )

    cfg: dict[str, Any] = {}
    for _, row in df.iterrows():
        field = str(row["field"]).strip()
        value = row["value"]
        if pd.isna(value):
            continue
        if isinstance(value, str):
            value = value.strip()
        # Coerce numerics where appropriate
        if field in {
            "occupancy", "floors", "zip_code",
            "supply_temp_f", "storage_temp_f", "inlet_temp_f",
        }:
            try:
                value = float(value) if "." in str(value) else int(value)
            except (TypeError, ValueError):
                pass
        cfg[field] = value
    cfg["__path__"] = str(p)
    return cfg


def list_buildings() -> list[str]:
    if not BUILDING_DIR.exists():
        return []
    return sorted(
        p.stem for p in BUILDING_DIR.glob("*.xlsx")
        if not p.name.startswith("_") and not p.name.startswith("~$")
    )


# ---------- meter-derived gpdpp ----------

def gpdpp_from_meter(meter_csv: str | Path, occupancy: float) -> float:
    """
    Compute gallons-per-day-per-person from a meter CSV.

    Reuses door-1-meters/meter.py if importable; otherwise does a lightweight
    inline computation (assumes gpm, auto-detects timestamp + flow columns).
    """
    if occupancy <= 0:
        raise ValueError("occupancy must be > 0 to derive gpdpp")

    meter_py = REPO_ROOT / "door-1-meters"
    if str(meter_py) not in sys.path:
        sys.path.insert(0, str(meter_py))
    try:
        from meter import load_meter, daily_totals  # type: ignore
        df = load_meter(meter_csv)
        daily = daily_totals(df)
        if daily.empty:
            raise ValueError("meter CSV produced no daily totals")
        avg_gpd = float(daily["gallons_per_day"].mean())
    except ImportError:
        # Fallback: inline minimal reader
        df = pd.read_csv(meter_csv)
        cols = [c.lower() for c in df.columns]
        df.columns = cols
        time_col = next((c for c in cols if "time" in c or "date" in c), cols[0])
        flow_col = next((c for c in cols if c != time_col and pd.api.types.is_numeric_dtype(df[c])), cols[1])
        df[time_col] = pd.to_datetime(df[time_col], errors="coerce")
        df = df.dropna(subset=[time_col]).sort_values(time_col)
        dt_min = df[time_col].diff().dt.total_seconds().median() / 60.0 or 5.0
        df["gallons"] = pd.to_numeric(df[flow_col], errors="coerce").fillna(0) * dt_min
        daily = df.groupby(df[time_col].dt.date)["gallons"].sum()
        avg_gpd = float(daily.mean())

    return avg_gpd / float(occupancy)


# ---------- ecosizer driver ----------

# Sensible defaults for a parallel_loop multi-family DHW system.
# These mirror the example in ecoengine's README and can be overridden
# via building config fields.
DEFAULT_ECOSIZER_KWARGS: dict[str, Any] = {
    "gpdpp":            25.0,
    "max_daily_run_hr": 16.0,
    "aquastat_fract":   0.4,
    "off_sensor_fract": 0.2,
    "on_trigger_t_f":   120.0,
    "off_trigger_t_f":  140.0,
    "return_flow_gpm":  3.0,
    "return_temp_f":    110.0,
    "tm_on_temp_f":     115.0,
    "tm_off_temp_f":    125.0,
    "tm_off_time_hr":   0.5,
}


def _assert_ecoengine() -> None:
    if not _ECOENGINE_OK:
        raise RuntimeError(
            f"ecoengine is not available: {_ECOENGINE_ERR}\n"
            f"Install with: pip install ecoengine"
        )


def build_ecosizer_kwargs(cfg: dict[str, Any], gpdpp: float | None = None) -> dict[str, Any]:
    """Translate a building-list config dict into EcosizerEngine kwargs."""
    kwargs = dict(DEFAULT_ECOSIZER_KWARGS)

    occupancy = cfg.get("occupancy")
    if occupancy is None:
        raise ValueError("building config missing 'occupancy'")

    kwargs.update({
        "building_type":            cfg.get("building_type", "multi_family"),
        "magnitude":                int(occupancy),
        "zip_code_or_climate_zone": int(cfg.get("zip_code", 94105)),
        "supply_temp_f":            float(cfg.get("supply_temp_f", 120.0)),
        "storage_temp_f":           float(cfg.get("storage_temp_f", 150.0)),
        "schematic":                cfg.get("schematic", "parallel_loop"),
    })

    if gpdpp is not None and not math.isnan(gpdpp):
        kwargs["gpdpp"] = float(gpdpp)

    # inlet_temp_f is honored only when the user also wants it; ecosizer
    # otherwise derives inlet from climate zone. We pass it via the
    # zip_code_or_climate_zone dict form when inlet override is desired.
    inlet = cfg.get("inlet_temp_f")
    if inlet is not None:
        kwargs["zip_code_or_climate_zone"] = {
            "design_oat_f": 35.0,  # placeholder; ecosizer may ignore if climate-looked-up
            "design_inlet_water_temp_f": float(inlet),
        }
    return kwargs


def run_ecosizer(
    cfg: dict[str, Any],
    meter_csv: str | Path | None = None,
    simulate: bool = False,
    simulate_out: str | Path | None = None,
) -> dict[str, Any]:
    """
    Run ecosizer for a building. Returns dict with 'sizing' and optionally
    'simulation_html_path'.
    """
    _assert_ecoengine()

    gpdpp = None
    if meter_csv:
        gpdpp = gpdpp_from_meter(meter_csv, float(cfg.get("occupancy", 0) or 0))

    kwargs = build_ecosizer_kwargs(cfg, gpdpp=gpdpp)

    engine = EcosizerEngine(**kwargs)
    engine.build()
    sizing = engine.get_sizing_results()

    out: dict[str, Any] = {
        "building": cfg.get("name") or Path(cfg["__path__"]).stem,
        "kwargs":   {k: v for k, v in kwargs.items() if not isinstance(v, dict)},
        "sizing":   sizing,
    }
    if gpdpp is not None:
        out["gpdpp_from_meter"] = round(gpdpp, 2)

    if simulate:
        result = engine.simulate_3day()
        html_path = str(simulate_out or "simulation.html")
        fig = result.to_plotly(
            title="3-Day Simulation",
            include_temperatures=True,
        )
        fig.write_html(html_path)
        out["simulation_html_path"] = html_path

    return out


# ---------- CLI ----------

def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="sizing-core — ecosizer wrapper.")
    ap.add_argument("--building", help="Building name (from building-list/) or .xlsx path.")
    ap.add_argument("--meter", default=None, help="Meter CSV for gpdpp derivation.")
    ap.add_argument("--simulate", action="store_true", help="Also run a 3-day design-day simulation.")
    ap.add_argument("--out", default="simulation.html", help="Path for simulation HTML output.")
    ap.add_argument("--list-buildings", action="store_true", help="List available buildings and exit.")
    args = ap.parse_args(argv)

    if args.list_buildings:
        names = list_buildings()
        if not names:
            print(f"(no buildings in {BUILDING_DIR})")
            return 0
        for n in names:
            print(n)
        return 0

    if not args.building:
        ap.error("--building is required (or use --list-buildings)")

    try:
        cfg = load_building(args.building)
    except (FileNotFoundError, ValueError) as e:
        sys.stderr.write(f"error: {e}\n")
        return 2

    try:
        result = run_ecosizer(
            cfg,
            meter_csv=args.meter,
            simulate=args.simulate,
            simulate_out=args.out if args.simulate else None,
        )
    except Exception as e:
        sys.stderr.write(f"ecosizer error: {e}\n")
        return 3

    print(f"building: {result['building']}")
    if "gpdpp_from_meter" in result:
        print(f"gpdpp (from meter): {result['gpdpp_from_meter']}")
    print("sizing:")
    for k, v in result["sizing"].items():
        print(f"  {k}: {v}")
    if "simulation_html_path" in result:
        print(f"simulation: {result['simulation_html_path']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
