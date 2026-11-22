"""
run_all.py — Run the whole pipeline end-to-end for one building + meter CSV.

Defaults: test_building_655 + the sample meter CSV in door-1-meters/.

Usage:
    python run_all.py
    python run_all.py --building test_building_655 --meter path/to/meter.csv
    python run_all.py --simulate            # also writes a 3-day sim.html
    python run_all.py --models options/models.csv   # also builds viable-models xlsx

Steps:
    1. meter summary (door-1-meters/meter.py)
    2. fixtures demo (door-2-fixtures/fixtures.py)        -- optional
    3. ecosizer sizing (sizing-core/size.py)
    4. building_results.xlsx (outputs/building_results.py)
    5. heat_pump_results.xlsx (outputs/heat_pump_results.py) -- only if --models
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent

DEFAULT_METER = REPO / "door-1-meters" / "Example DHW Flow Trend Hourly & Daily(LOG01A (RAW DATA)).csv"
FIXTURES_XLSX = REPO / "door-2-fixtures" / "general_purpose_hot_water_requirements.xlsx"


def _import(folder: str, module: str):
    target = REPO / folder
    if str(target) not in sys.path:
        sys.path.insert(0, str(target))
    return __import__(module)


def _banner(n: int, title: str) -> None:
    print()
    print("=" * 60)
    print(f"[{n}] {title}")
    print("=" * 60)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Run the full pipeline.")
    ap.add_argument("--building", default="test_building_655", help="Building name from building-list/.")
    ap.add_argument("--meter",    default=str(DEFAULT_METER), help="Meter CSV path.")
    ap.add_argument("--simulate", action="store_true", help="Also run a 3-day design-day simulation.")
    ap.add_argument("--models",   default=None, help="Optional models CSV/XLSX for heat pump viability.")
    ap.add_argument("--out-dir",  default=str(REPO / "outputs"), help="Where to put generated xlsx files.")
    ap.add_argument("--skip-fixtures", action="store_true", help="Skip the fixtures demo step.")
    args = ap.parse_args(argv)

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    meter = Path(args.meter)
    if not meter.exists():
        sys.stderr.write(f"error: meter CSV not found: {meter}\n")
        return 2

    # --- 1. meter summary ---
    _banner(1, f"meter summary ({meter.name})")
    meter_mod = _import("door-1-meters", "meter")
    df = meter_mod.load_meter(meter)
    summary = meter_mod.summary(df)
    width = max(len(k) for k in summary)
    for k, v in summary.items():
        print(f"  {k.ljust(width)}  {v}")

    # --- 2. fixtures demo ---
    if not args.skip_fixtures and FIXTURES_XLSX.exists():
        _banner(2, "fixtures demo (2 Bar sinks + 4 Lavatories + 1 Triple comp)")
        fixtures_mod = _import("door-2-fixtures", "fixtures")
        fdf = fixtures_mod.load_fixtures(FIXTURES_XLSX)
        totals = fixtures_mod.total_demand(
            fdf,
            {"Bar sink": 2, "Lavatories": 4, "Triple compartment sink": 1},
        )
        for k, v in totals.items():
            print(f"  {k}: {v}")

    # --- 3. sizing ---
    _banner(3, f"sizing ({args.building})")
    size_mod = _import("sizing-core", "size")
    cfg = size_mod.load_building(args.building)
    sim_html = str((out_dir / "simulation.html").resolve()) if args.simulate else None
    result = size_mod.run_ecosizer(
        cfg,
        meter_csv=meter,
        simulate=args.simulate,
        simulate_out=sim_html,
    )
    print(f"  building:            {result['building']}")
    if "gpdpp_from_meter" in result:
        print(f"  gpdpp (from meter):  {result['gpdpp_from_meter']}")
    for k, v in result["sizing"].items():
        print(f"  {k}: {v}")
    if "simulation_html_path" in result:
        print(f"  simulation html:     {result['simulation_html_path']}")

    # --- 4. building_results.xlsx ---
    _banner(4, "building_results.xlsx")
    br_mod = _import("outputs", "building_results")
    br_out = out_dir / "building_results.xlsx"
    br_path = br_mod.export_building_results(
        meter_csv=meter,
        out=br_out,
        building_name=args.building,
    )
    print(f"  wrote {br_path}")

    # --- 5. heat_pump_results.xlsx (only if models provided) ---
    if args.models:
        _banner(5, "heat_pump_results.xlsx")
        hp_mod = _import("outputs", "heat_pump_results")
        hp_out = out_dir / "heat_pump_results.xlsx"
        try:
            hp_path = hp_mod.export_heat_pump_results(
                building=args.building,
                models_path=args.models,
                out=hp_out,
                meter_csv=meter,
            )
            print(f"  wrote {hp_path}")
        except (FileNotFoundError, ValueError) as e:
            print(f"  skipped: {e}")
    else:
        _banner(5, "heat_pump_results.xlsx (skipped — no --models provided)")

    print()
    print("done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
