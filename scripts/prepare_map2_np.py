"""Map 2 national-parks-only data: python prepare_map2_np.py (run after prepare_data.py).

Overwrites capad2024_terrestrial_jurisdiction.csv in place, so map2-choropleth.json
needs no change. Land area (the denominator) is kept from the official CAPAD
'Jurisdiction' sheet that prepare_data.py wrote; only the protected area,
percentage and count are replaced with national-park-only figures.

Why these choices:
- "National park" = the CAPAD TYPE values that name a park outright. Proposed
  parks are not gazetted yet. Victoria's Schedule 4/6 areas sit inside existing
  parks, so counting them would count the same land twice.
- OVERLAP == "2" marks a record that overlaps another record, so it is skipped
  to avoid double counting.
- Area is GIS_AREA (hectares from the polygon). The official sheet already
  allows for overlaps; this sum is close to it but not identical.
"""
import csv
import json
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "raw-data" / "Collaborative_Australian_Protected_Areas_Database_(CAPAD)_2024_-_Terrestrial__.json"
OUTPUT = ROOT / "docs" / "data" / "capad2024_terrestrial_jurisdiction.csv"
NP_TYPES = {
    "National Park", "National Park (Commonwealth)", "National Park Aboriginal",
    "CCA Zone 1 National Park", "National Park (Scientific)",
}
STATE_NAMES = {
    "ACT": "Australian Capital Territory", "JBT": "Jervis Bay Territory",
    "NSW": "New South Wales", "NT": "Northern Territory", "QLD": "Queensland",
    "SA": "South Australia", "TAS": "Tasmania", "VIC": "Victoria", "WA": "Western Australia",
}

# Read the land areas already in the csv before overwriting it.
with open(OUTPUT, encoding="utf-8") as f:
    land = {row["code"]: float(row["jurisdiction_area_ha"]) for row in csv.DictReader(f)}

topo = json.loads(SOURCE.read_text(encoding="utf-8"))
area, ids = defaultdict(float), defaultdict(set)
for geometry in next(iter(topo["objects"].values()))["geometries"]:
    p = geometry["properties"]
    name = STATE_NAMES.get(p["STATE"])   # skips EXT (external territories)
    if p["TYPE"] in NP_TYPES and p["OVERLAP"] != "2" and name:
        area[name] += p["GIS_AREA"] or 0
        ids[name].add(p["PA_ID"])

with open(OUTPUT, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["code", "jurisdiction_area_ha", "protected_area_ha", "pct_protected", "number_of_pas"])
    for name in sorted(land):
        writer.writerow([name, round(land[name]), round(area[name]),
                         round(100 * area[name] / land[name], 2), len(ids[name])])
print(f"{OUTPUT.name}: national parks only, {len(land)} rows")
