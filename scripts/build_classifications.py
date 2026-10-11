"""
Make national_park_tourism_regions_verified.csv the ONE source of truth for "is this a
national-park region?", and rebuild every other classification file from it.

    national_park_tourism_regions_verified.csv          (you edit THIS file, and only this one)
      |
      +--> national_parks_tourism_regions.csv           (old 2-column format, same flags)
      +--> national_park_routes.csv                     (route = Yes if EITHER end is a Yes region)
      +--> national_park_airport_classification.csv     (airport = its region's flag)

Run it with:  python build_classifications.py
then re-run the chart scripts:  prepare_vis67.py, prepare_vis89.py, prepare_vis10.py

WHY THIS EXISTS (2026-10-08)
The four csv files had drifted apart. The older region list said No to Peninsula, Central Coast
and Lakes, which the verified list says Yes to. The route and airport files were built from the
older list, and marked 17 airports' regions "Not in earlier region list" (Darwin, the
Whitsundays, the Pilbara...). So vis 6/7, vis 8/9 and vis 10 were each using a different
definition of a national-park region. Every file is now derived from the verified list by the
same rules, so they cannot disagree again unless one is edited by hand.

WHAT IS A JUDGEMENT HERE: only the AIRPORT -> REGION table below. Each airport is placed in the
tourism region it physically sits in. The non-obvious ones carry a note.
"""

from pathlib import Path

# ---- Folder layout (this script lives in scripts/, one level inside the project) ----
ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "raw-data"                 # source files exactly as downloaded
CLASS_DIR = ROOT / "classifications"        # the hand-verified Yes/No region, route and airport csvs
DATA_DIR = ROOT / "docs" / "data"           # processed data the charts load (inside the website)

import csv
import io

VERIFIED = CLASS_DIR / "national_park_tourism_regions_verified.csv"
OUT_REGIONS = CLASS_DIR / "national_parks_tourism_regions.csv"
OUT_ROUTES = CLASS_DIR / "national_park_routes.csv"
OUT_AIRPORTS = CLASS_DIR / "national_park_airport_classification.csv"

# Airport code -> (city, tourism region as spelled in the verified csv, note). Codes are the ones
# BITRE uses in the Top Routes sheet. A region name that is not in the verified csv stops the run.
AIRPORTS = {
    "ABX": ("Albury", "The Murray", "Albury is in the NSW Murray tourism region."),
    "ADL": ("Adelaide", "Adelaide", ""),
    "ARM": ("Armidale", "New England North West", ""),
    "ASP": ("Alice Springs", "Alice Springs", "The town region, separate from MacDonnell (verified csv)."),
    "AYQ": ("Ayers Rock / Uluru", "Lasseter", "Connellan Airport, Yulara - the Uluru-Kata Tjuta gateway."),
    "BDB": ("Bundaberg", "Bundaberg", ""),
    "BME": ("Broome", "Australia's North West", "Kimberley, inside the Australia's North West region."),
    "BNE": ("Brisbane", "Brisbane", ""),
    "CBR": ("Canberra", "Canberra", ""),
    "CFS": ("Coffs Harbour", "North Coast NSW", ""),
    "CNS": ("Cairns", "Tropical North Queensland", ""),
    "DBO": ("Dubbo", "Central NSW", ""),
    "DPO": ("Devonport", "North West", "Tasmania's North West, not WA's North West."),
    "DRW": ("Darwin", "Darwin", "City region; Kakadu and Litchfield are a separate region (verified csv)."),
    "EMD": ("Emerald", "Capricorn", "Central Highlands (Qld) is part of the wider Capricorn region."),
    "GET": ("Geraldton", "Australia's Coral Coast", ""),
    "GLT": ("Gladstone", "Gladstone", ""),
    "HBA": ("Hobart", "Hobart and the South", ""),
    "HTI": ("Hamilton Island", "Whitsundays", ""),
    "ISA": ("Mount Isa", "Outback Queensland", ""),
    "KGI": ("Kalgoorlie", "Australia's Golden Outback", ""),
    "KTA": ("Karratha", "Australia's North West", "Pilbara, inside the Australia's North West region."),
    "LST": ("Launceston", "Launceston and the North", ""),
    "MCY": ("Sunshine Coast", "Sunshine Coast", ""),
    "MEL": ("Melbourne", "Melbourne", ""),
    "MKY": ("Mackay", "Mackay", ""),
    "MOV": ("Moranbah", "Mackay", "Isaac region, inside the Mackay tourism region."),
    # Previously mapped to "The Murray". The verified csv says The Murray is the NSW region and
    # lists Victoria's Murray subregions separately; Mildura is the centre of the Mallee region.
    "MQL": ("Mildura", "Mallee", "Corrected from The Murray (NSW) - Mildura is in Victoria's Mallee."),
    "NTL": ("Newcastle", "Hunter", ""),
    "OAG": ("Orange", "Central NSW", ""),
    "OOL": ("Gold Coast", "Gold Coast", ""),
    "PER": ("Perth", "Destination Perth", ""),
    "PHE": ("Port Hedland", "Australia's North West", "Pilbara, inside the Australia's North West region."),
    "PLO": ("Port Lincoln", "Eyre Peninsula", ""),
    "PPP": ("Proserpine / Whitsunday Coast", "Whitsundays", ""),
    "PQQ": ("Port Macquarie", "North Coast NSW", ""),
    "QNA": ("Ballina", "North Coast NSW", "BITRE's own route names print QNA as Ballina."),
    "ROK": ("Rockhampton", "Capricorn", ""),
    "SYD": ("Sydney", "Sydney", ""),
    "TMW": ("Tamworth", "New England North West", ""),
    "TSV": ("Townsville", "Townsville", ""),
    "WGA": ("Wagga Wagga", "Riverina", ""),
    "ZNE": ("Newman", "Australia's North West", "Pilbara, inside the Australia's North West region."),
}

# The airport names used by the "Airport Passengers" sheet (vis 9), mapped to the codes above.
AIRPORT_NAMES = {
    "ADELAIDE": "ADL", "BALLINA": "QNA", "BRISBANE": "BNE", "CAIRNS": "CNS", "CANBERRA": "CBR",
    "DARWIN": "DRW", "GOLD COAST": "OOL", "HAMILTON ISLAND": "HTI", "HOBART": "HBA",
    "KARRATHA": "KTA", "LAUNCESTON": "LST", "MACKAY": "MKY", "MELBOURNE": "MEL",
    "NEWCASTLE": "NTL", "PERTH": "PER", "PROSERPINE": "PPP", "ROCKHAMPTON": "ROK",
    "SUNSHINE COAST": "MCY", "SYDNEY": "SYD", "TOWNSVILLE": "TSV",
}

ROUTE_RULE = ("Verified region list; either-endpoint rule; geographic proxy, "
              "not evidence of passenger purpose.")


def read_csv(path):
    with io.open(path, "r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(path, rows, fields):
    with io.open(path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
    print("wrote {0}  ({1} rows)".format(path, len(rows)))


# ---- 1. The source of truth ---------------------------------------------------
verified = read_csv(VERIFIED)
flag = {r["tourism_region"].strip(): r["national_park_region"].strip() for r in verified}
bad = {f for f in flag.values()} - {"Yes", "No"}
if bad:
    raise SystemExit("Unexpected flag values in {0}: {1}".format(VERIFIED, bad))
unknown = sorted({reg for _, reg, _ in AIRPORTS.values()} - set(flag))
if unknown:
    raise SystemExit("Airport regions not in {0}: {1}".format(VERIFIED, unknown))

# ---- 2. The old two-column region list, now just a copy of the verified flags -----
# Kept because it is the format the earlier charts were written against; nothing reads it now.
write_csv(OUT_REGIONS,
          [{"Region": r["tourism_region"], "Popular for national parks tourism": r["national_park_region"]}
           for r in verified],
          ["Region", "Popular for national parks tourism"])

# ---- 3. Routes: same rows and columns as before, flags re-derived ------------------
# Every existing column is kept (source_row_count, mapping_source_urls...), so nothing you added
# by hand is lost; only the region and flag columns are recomputed.
routes = read_csv(OUT_ROUTES)
fields = list(routes[0].keys())
changes = []
for r in routes:
    old = r["national_park_route"]
    for end in ("1", "2"):
        code = r["airport_" + end]
        if code not in AIRPORTS:
            raise SystemExit("Airport {0} (route {1}) is not in AIRPORTS".format(code, r["route_pair"]))
        city, region, _ = AIRPORTS[code]
        r["airport_{0}_city".format(end)] = city
        r["airport_{0}_region".format(end)] = region
        r["airport_{0}_region_flag".format(end)] = flag[region]
    yes_end = flag[AIRPORTS[r["airport_1"]][1]] == "Yes" or flag[AIRPORTS[r["airport_2"]][1]] == "Yes"
    r["national_park_route"] = "Yes" if yes_end else "No"
    r["classification_reason"] = ("At least one endpoint is in a national-park region." if yes_end
                                  else "Neither endpoint is in a national-park region.")
    notes = [AIRPORTS[r["airport_" + e]][2] for e in ("1", "2") if AIRPORTS[r["airport_" + e]][2]]
    r["mapping_notes"] = " ".join(notes)
    r["classification_basis"] = ROUTE_RULE
    if old != r["national_park_route"]:
        changes.append("{0}: {1} -> {2}".format(r["route_pair"], old, r["national_park_route"]))
write_csv(OUT_ROUTES, routes, fields)
print("  route flags changed: " + ("; ".join(changes) if changes else "none"))

# ---- 4. Airports: ONE row per airport ----------------------------------------------
# The old file had one row per sheet row (4,253) with a source_row column that no longer lined
# up with the sheet (see prepare_vis89.py, trap 5). The flag belongs to the AIRPORT, not to a
# month, so one row each is both correct and easy to check by eye.
airports = []
for name, code in AIRPORT_NAMES.items():
    city, region, note = AIRPORTS[code]
    airports.append({"airport": name, "iata_code": code, "tourism_region": region,
                     "national_park_tourism": flag[region], "classification_note": note})
airports.append({"airport": "All Australian Airports", "iata_code": "", "tourism_region": "Aggregate",
                 "national_park_tourism": "Not applicable",
                 "classification_note": "Aggregate of all airports; not a destination."})
write_csv(OUT_AIRPORTS, airports, list(airports[0].keys()))
print("  Yes airports: " + ", ".join(a["airport"] for a in airports if a["national_park_tourism"] == "Yes"))
