"""
Turn the June 2024 tables of the Tourism Businesses in Australia workbook into ONE tidy csv for
vis 10 (butterfly chart of businesses by employment size).

    tra-tbia-2025-d-geographic-breakdown.xlsx, the eight STATE sheets, June 2024 table on each
    national_park_tourism_regions_verified.csv                  (your verified Yes/No per region)
      |
      +--> data/np_business_size.csv   (vis 10, 1 row per side x city-or-not x size band)

Run it with:  python prepare_vis10.py
Needs openpyxl (pip install openpyxl) because the source is .xlsx.

DECISIONS (Nicholas, 2026-10-08)
  • LEFT side  = every region the verified csv marks "Yes"  (53 regions).
  • RIGHT side = every region the verified csv marks "No"   (23 regions).
  • Each side is split into CAPITAL-CITY regions and the rest, so the chart can show the cities in
    a darker shade. Hobart and the South and Canberra are capitals the csv marks Yes, so the left
    side has a (small) city part too - leaving it out would suggest the park side has no cities.
  • Measure = "Tourism Industries Total": tourism-characteristic PLUS tourism-connected
    businesses, which is TRA's headline count of "tourism businesses".
  • Year = June 2024 (your stub: "to align with the tourism and CAPAD reference year").

--------------------------------------------------------------------------------------------
TRAPS IN THIS WORKBOOK - read before changing anything
--------------------------------------------------------------------------------------------
1.  EACH STATE SHEET STACKS FIVE TABLES, one per year (June 2021 ... June 2025), one under the
    other. Row numbers differ by state, so the script FINDS the June 2024 table by its title
    ("... BY EMPLOYMENT SIZE, JUNE 2024") instead of using fixed row numbers.

2.  THE AUSTRALIA SHEET IS JUNE 2025 ONLY. That is why the state sheets are read, as your stub
    says - the national sheet has no June 2024 table at all.

3.  EACH REGION HAS THREE ROWS: "Tourism Characteristic Total", "Tourism Connected Total" and
    "Tourism Industries Total" (the sum of the other two). Only the last is read. The ACT table
    is the exception: its third row is called "Total tourism-related", so both labels are
    accepted.

4.  THE REGION NAME IS ONLY ON THE FIRST OF ITS THREE ROWS; the other two have a blank or a
    single space in column A. The script carries the last name it saw down to them.

5.  THE CAPITAL-CITY REGION COMES FIRST, WITH NO GAP. In every state table the capital (e.g.
    Sydney) is on the line IMMEDIATELY after the "Tourism region" header line. Skipping even one
    row too many silently drops the biggest region in each state - which happened on the first
    exploratory read of this sheet. The script therefore starts reading straight after the
    table's title and simply ignores the header lines.

6.  AGGREGATE ROWS. Each table ends with "<State> Regional Areas", "<State> Capital City" and
    "<State> Total". They are sums of the regions above them and are excluded. Queensland's is
    "Capital City & Gold Coast" - the SOURCE lumps the Gold Coast in with Brisbane there.

7.  ROUNDING. The ABS perturbs small cells (many 0s and 3s), and the sheet's own footnote says
    the size bands "may not equal to totals due to rounding". So the chart's percentages are
    computed from the SUM OF THE FIVE BANDS (they then add to exactly 100%), while the total
    printed on each side is the sum of the sheet's own Grand Totals. The script reports how far
    apart the two are (a handful of businesses out of 100,000+).

NO MISSING DATA. All 76 regions in the verified csv are found in the June 2024 tables, and every
cell is a number (no "np", no blanks). The script stops if either stops being true.
"""

from pathlib import Path

# ---- Folder layout (this script lives in scripts/, one level inside the project) ----
ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "raw-data"                 # source files exactly as downloaded
CLASS_DIR = ROOT / "classifications"        # the hand-verified Yes/No region, route and airport csvs
DATA_DIR = ROOT / "docs" / "data"           # processed data the charts load (inside the website)

import csv           # reads your classification file and writes the output
import io            # lets us force the encoding on the way in and out
import os            # makes sure the data/ folder exists before writing into it

try:
    import openpyxl  # the only way to read .xlsx; not in the standard library
except ImportError:
    raise SystemExit("This script needs openpyxl.  Run:  pip install openpyxl")

# ---- 1. Settings you may want to change -------------------------------------

IN_XLSX = RAW_DIR / "tra-tbia-2025-d-geographic-breakdown.xlsx"
IN_FLAGS = CLASS_DIR / "national_park_tourism_regions_verified.csv"
OUT_CSV = os.path.join(DATA_DIR, "np_business_size.csv")

# The eight state/territory sheets. The AUSTRALIA sheet is skipped on purpose - see trap 2.
STATE_SHEETS = ["NSW", "Vic", "Qld", "SA", "WA", "Tas", "NT", "ACT"]

# Which year's table to read, matched against the END of each table title.
TABLE_TITLE_SUFFIX = "BY EMPLOYMENT SIZE, JUNE 2024"

# The row label that holds the figure we want - see trap 3.
MEASURE_LABELS = {"Tourism Industries Total", "Total tourism-related"}

# Side names. These are what the chart's tooltip prints; the chart's headings and legend labels
# are written separately in vis10-butterfly.js.
SIDE_PARK = "park"
SIDE_COMPARISON = "comparison"

# The capital-city regions. Written out rather than inferred, then CHECKED against the sheet's
# own "<State> Capital City" rows in step 5, so the definition is the source's, not ours.
# Gold Coast is NOT here even though Queensland's aggregate row includes it (trap 6): it is not
# a capital city. That check is done for Queensland with Gold Coast added back in.
CAPITAL_REGIONS = {
    "NSW": "Sydney", "Vic": "Melbourne", "Qld": "Brisbane", "SA": "Adelaide",
    "WA": "Destination Perth", "Tas": "Hobart and the South", "NT": "Darwin", "ACT": "Canberra",
}

# The five employment-size bands, in the order the chart draws them (smallest business at the
# top). Each entry: (label written to the csv, the 0-based position of that band's number in
# the row's figures - see COLUMNS below).
#   Columns in every June 2024 table (1-based, as Excel counts):
#     C Micro (1-4) | D Small (5-19) | E Medium (20-199) | F Large (200+) |
#     G Total employing | H Non employing | I Grand Total
FIRST_FIGURE_COLUMN = 3          # column C - the figures are read as row[2:9]
BANDS = [
    ("Non-employing", 5),         # column H. Sole traders etc. - nobody on the payroll.
    ("Micro (1–4)", 0),      # column C
    ("Small (5–19)", 1),     # column D
    ("Medium (20–199)", 2),  # column E
    ("Large (200+)", 3),          # column F
]
GRAND_TOTAL_INDEX = 6             # column I


def is_number(value):
    """True for a real count. bool is excluded because Python treats True as 1."""
    return isinstance(value, (int, float)) and not isinstance(value, bool)


# ---- 2. Read the June 2024 table from every state sheet ----------------------
book = openpyxl.load_workbook(IN_XLSX, read_only=True, data_only=True)

regions = {}          # "Sydney" -> (state sheet, [7 figures from columns C-I])
aggregates = {}       # "NSW Capital City" -> [7 figures], used for the capital-city check

for state in STATE_SHEETS:
    rows = list(book[state].iter_rows(values_only=True))
    # Trap 1: find the June 2024 table by its title.
    start = next((i for i, r in enumerate(rows)
                  if isinstance(r[0], str) and r[0].strip().endswith(TABLE_TITLE_SUFFIX)), None)
    if start is None:
        raise SystemExit("No June 2024 table found on sheet {0!r}".format(state))

    current = None                                   # the region name being carried down (trap 4)
    for row in rows[start + 1:]:
        first = row[0].strip() if isinstance(row[0], str) else row[0]
        # The table ends at its footnote or at the next year's title.
        if isinstance(first, str) and (first.startswith("Note") or first.upper().startswith("TABLE")):
            break
        # A non-blank column A is a new region name. "Tourism region" is header text, not a
        # region, so it is ignored rather than carried down. The capital's name is on the very
        # next line (trap 5). The other header line has only blanks or a single space in column
        # A, which .strip() turns into "", so it is ignored too.
        if first and first != "Tourism region":
            current = first
        label = row[1].strip() if isinstance(row[1], str) else row[1]
        if label not in MEASURE_LABELS:
            continue                                 # one of the other two measure rows (trap 3)
        figures = list(row[FIRST_FIGURE_COLUMN - 1:FIRST_FIGURE_COLUMN - 1 + 7])
        if not all(is_number(v) for v in figures):
            raise SystemExit("Non-numeric cell for {0} / {1}: {2}".format(state, current, figures))
        # Trap 6: aggregate rows are kept aside, never treated as regions.
        if current.startswith(state + " ") and any(
                current.endswith(s) for s in ("Regional Areas", "Capital City", "Capital City & Gold Coast", "Total")):
            aggregates[current] = figures
        else:
            if current in regions:
                raise SystemExit("Region {0!r} appears twice".format(current))
            regions[current] = (state, figures)

print("read {0} regions from the June 2024 tables of {1} sheets".format(len(regions), len(STATE_SHEETS)))

# ---- 3. Read your verified classification --------------------------------------
# utf-8-sig so a byte-order mark from Excel cannot glue itself to the first header name.
flags = {}
with io.open(IN_FLAGS, "r", encoding="utf-8-sig", newline="") as handle:
    for row in csv.DictReader(handle):
        flags[row["tourism_region"].strip()] = row["national_park_region"].strip()

# Every classified region must be in the sheet, and vice versa. A name that fails to match would
# otherwise silently drop that region's businesses from its side.
missing = sorted(set(flags) - set(regions))
unclassified = sorted(set(regions) - set(flags))
if missing or unclassified:
    raise SystemExit("Name mismatch.\n  in csv, not in sheet: {0}\n  in sheet, not in csv: {1}"
                     .format(missing, unclassified))
unknown_flags = {f for f in flags.values()} - {"Yes", "No"}
if unknown_flags:
    raise SystemExit("Unexpected values in national_park_region: {0}".format(unknown_flags))
print("read {0} classifications from {1!r}: {2} Yes, {3} No - all match the sheet".format(
    len(flags), IN_FLAGS, sum(f == "Yes" for f in flags.values()), sum(f == "No" for f in flags.values())))

# ---- 4. Check the capital-city list against the source's own aggregate rows -----
for state, capital in CAPITAL_REGIONS.items():
    if state == "ACT":
        continue        # the ACT table has a single region and no aggregate rows to compare with
    if state == "Qld":
        # Trap 6: the source's Queensland row is "Capital City & Gold Coast".
        key, ours = "Qld Capital City & Gold Coast", [a + b for a, b in zip(
            regions["Brisbane"][1], regions["Gold Coast"][1])]
    else:
        key, ours = state + " Capital City", regions[capital][1]
    # Allow a few businesses of difference for the ABS rounding (trap 7).
    if any(abs(a - b) > 6 for a, b in zip(ours, aggregates[key])):
        raise SystemExit("{0} does not match the sheet's {1!r} row:\n  {2}\n  {3}".format(
            capital, key, ours, aggregates[key]))
print("capital-city regions confirmed against each state's 'Capital City' row")
capitals = set(CAPITAL_REGIONS.values())

# ---- 5. Sum each side x city-or-not x size band ----------------------------------
sides = {SIDE_PARK: "Yes", SIDE_COMPARISON: "No"}
out = []
report = []
for side, flag in sides.items():
    members = sorted(r for r, f in flags.items() if f == flag)
    # Denominator for the percentages: the SUM OF THE FIVE BANDS across the whole side (trap 7).
    band_sum = sum(regions[r][1][i] for r in members for _, i in BANDS)
    grand_total = sum(regions[r][1][GRAND_TOTAL_INDEX] for r in members)
    report.append((side, len(members), grand_total, band_sum))
    for is_capital in (True, False):
        part = [r for r in members if (r in capitals) == is_capital]
        for order, (band, i) in enumerate(BANDS, 1):
            businesses = sum(regions[r][1][i] for r in part)
            out.append({
                "side": side,
                "is_capital": 1 if is_capital else 0,     # 1/0, not True/False - see vis 9's note
                "band": band,
                "band_order": order,
                "businesses": businesses,
                # A FRACTION of the whole side, so the city and non-city parts of one band add up
                # to that band's share, and the five bands add up to 1.
                "share": round(businesses / band_sum, 6),
                "side_total": grand_total,                # printed under each side's heading
                "side_regions": len(members),
                "part_regions": len(part),
                # The regions behind this bar segment, for the tooltip. Semicolons rather than
                # commas so the csv stays easy to read in Excel.
                "region_names": "; ".join(part),
            })

os.makedirs("data", exist_ok=True)
with io.open(OUT_CSV, "w", encoding="utf-8", newline="") as handle:
    writer = csv.DictWriter(handle, fieldnames=list(out[0].keys()))
    writer.writeheader()
    writer.writerows(out)
print("wrote {0}  ({1} rows)\n".format(OUT_CSV, len(out)))

# ---- 6. A short report to sanity-check against the chart -------------------------
for side, n, grand_total, band_sum in report:
    print("{0:<11} {1:>2} regions   grand total {2:>8,}   sum of bands {3:>8,}   (difference {4:+,})".format(
        side, n, grand_total, band_sum, band_sum - grand_total))
print("\nshare of each side's businesses (%), city part in brackets:")
for band, _ in BANDS:
    cells = []
    for side in sides:
        rows = [r for r in out if r["side"] == side and r["band"] == band]
        total = sum(r["share"] for r in rows) * 100
        city = sum(r["share"] for r in rows if r["is_capital"]) * 100
        cells.append("{0:5.1f} ({1:4.1f})".format(total, city))
    print("  {0:<18} park {1}    comparison {2}".format(band, *cells))
