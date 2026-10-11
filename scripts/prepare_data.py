"""
prepare_data.py – shrink the CAPAD 2024 files into two small CSVs for the website.

Run once (needs Python 3 + openpyxl:  pip install openpyxl):
    python scripts/prepare_data.py

Inputs  (in raw-data/):
    Collaborative_Australian_Protected_Areas_Database_(CAPAD)_2024_-_Terrestrial__.json
    Collaborative_Australian_Protected_Areas_Database_(CAPAD)_2024_-_Marine.json
    capad-2024-terrestrial-national.xlsx
Outputs (written straight into docs/data/, the website's data folder):
    capad2024_points.csv                     -> Map 1 (one row per CAPAD record)
    capad2024_terrestrial_jurisdiction.csv   -> Map 2 (official CAPAD 2024 table)
"""
import csv
import json
from datetime import datetime, timedelta

from pathlib import Path

# ---- Folder layout (this script lives in scripts/, one level inside the project) ----
ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "raw-data"                 # source files exactly as downloaded
CLASS_DIR = ROOT / "classifications"        # the hand-verified Yes/No region, route and airport csvs
DATA_DIR = ROOT / "docs" / "data"           # processed data the charts load (inside the website)

STATE_NAMES = {"NSW": "New South Wales", "VIC": "Victoria", "QLD": "Queensland", "SA": "South Australia",
               "WA": "Western Australia", "TAS": "Tasmania", "NT": "Northern Territory",
               "ACT": "Australian Capital Territory", "JBT": "Jervis Bay Territory",
               "COM": "Commonwealth", "EXT": "External Territories"}

import openpyxl

TERRESTRIAL = RAW_DIR / "Collaborative_Australian_Protected_Areas_Database_(CAPAD)_2024_-_Terrestrial__.json"
MARINE      = RAW_DIR / "Collaborative_Australian_Protected_Areas_Database_(CAPAD)_2024_-_Marine.json"
SUMMARY     = RAW_DIR / "capad-2024-terrestrial-national.xlsx"


def australian_year(gaz_date):
    """'1965-12-31T13:00:00Z' is midnight 1 Jan 1966 in Sydney -> return 1966.
    Adding 12 hours to the UTC time gives the Australian calendar date."""
    if not gaz_date:
        return ""                                   # empty = unknown date
    utc = datetime.strptime(gaz_date, "%Y-%m-%dT%H:%M:%SZ")
    return (utc + timedelta(hours=12)).year


def read_points(path, environment):
    """Read ONLY the attribute table of a TopoJSON file (the polygons are ignored)."""
    with open(path, encoding="utf-8") as f:
        topo = json.load(f)
    layer = next(iter(topo["objects"].values()))   # the file has one object
    rows = []
    for geometry in layer["geometries"]:
        p = geometry.get("properties") or {}
        rows.append({
            "name":     p["NAME"],
            "type":     p["TYPE"],
            "state":    STATE_NAMES.get(p["STATE"], p["STATE"]),
            "env":      environment,                       # T = terrestrial, M = marine
            "lon":      round(p["LONGITUDE"], 2),          # 0.01° ≈ 1 km (less than 1 pixel)
            "lat":      round(p["LATITUDE"], 2),
            "hectares": round(p["GIS_AREA"], 1),
            "year":     australian_year(p["GAZ_DATE"]),
        })
    return rows


# ---------- Map 1: points ----------
points = read_points(TERRESTRIAL, "T") + read_points(MARINE, "M")
with open(DATA_DIR / "capad2024_points.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=list(points[0].keys()))
    writer.writeheader()
    writer.writerows(points)
print(f"capad2024_points.csv: {len(points)} rows")

# ---------- Map 2: official jurisdiction table ----------
sheet = openpyxl.load_workbook(SUMMARY, data_only=True)["Jurisdiction"]
wanted = {"ACT", "JBT", "NSW", "NT", "QLD", "SA", "TAS", "VIC", "WA"}   # skips EXT and totals
with open(DATA_DIR / "capad2024_terrestrial_jurisdiction.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["code", "jurisdiction_area_ha", "protected_area_ha", "pct_protected", "number_of_pas"])
    for row in sheet.iter_rows(values_only=True):
        if row[0] in wanted:
            code, area, number, protected, pct = row[0], row[1], row[2], row[3], row[4]
            writer.writerow([STATE_NAMES.get(code, code), round(area), round(protected), round(pct, 2), number])
print("capad2024_terrestrial_jurisdiction.csv: 9 rows")
