"""
Turn Table 10 of the National Visitor Survey workbook into two TIDY csv files.

    tra-nvs-quarterly-december-2024-2023.xlsx, sheet "Table 10"
    national_park_tourism_regions_verified.csv    (your verified Yes/No decision per region)
      |
      +--> data/np_trips_pairs.csv      (vis 6, 1 row per selected region, both years side by side)
      +--> data/np_exp_per_trip.csv     (vis 7, 1 row per plotted region, park + capital groups)

Run it with:  python prepare_vis67.py
Needs openpyxl (pip install openpyxl) because the source is .xlsx, not .csv.

WHY A PREP SCRIPT AT ALL?
Same reason as prepare_vis45.py: keep the two .js specs about DESIGN, not cleaning. Vega-Lite
cannot read .xlsx at all, and this sheet needs four things done to it that are awkward in a spec
(see the TRAPS block below). Doing them here means the charts read a flat csv and the .js files
stay about colour, layout and sight lines.

--------------------------------------------------------------------------------------------
TRAPS IN THIS SHEET - read before changing anything
--------------------------------------------------------------------------------------------
1.  TABLE 16 IS NOT TABLE 10. The stub comments in vis6-bar.js and vis7-lolipop.js originally
    said "Table 16", but Table 16 is "DAY trips and regional expenditure for top 50 regions".
    Overnight trips for the top 50 regions is TABLE 10. Nicholas confirmed (2026-10-07) that
    overnight trips is the measure he wants, so this script reads Table 10.
    The two tables have DIFFERENT top-50 rosters - 8 regions each that the other lacks - so you
    cannot mix figures from them. If you ever switch back to Table 16, note that it has no
    "visitor nights" columns and that its column positions are therefore different.

2.  THE SHEET'S OWN RANK COLUMN IS BY NIGHTS, NOT TRIPS. Row 4 says "RANKED BY NIGHTS IN REGION
    FOR YEAR ENDING DECEMBER 2024". Using column A as a trips ranking would be wrong: Adelaide
    is 10th by nights but 10th by trips only by coincidence, while Tropical North Queensland is
    17th by trips and well above that by nights. This script therefore IGNORES column A and
    computes its own ranking from the 2024 trips figure.

3.  TRIPS ARE IN THOUSANDS, EXPENDITURE IN MILLIONS OF DOLLARS. The headers say Trips ('000) and
    Regional expenditure ($M). Mixing the two units is how you get an "average spend per trip"
    that is out by a factor of 1000. This script multiplies trips by 1,000 so the csv holds
    whole trips, which is what a reader expects on an axis.

4.  "np" MEANS NOT PUBLISHED. Table 16 is riddled with it (10 of 50 regions have no 2024
    expenditure). TABLE 10 HAS NONE - every one of its 50 regions has a complete figure in both
    years for trips, nights, expenditure and expenditure per trip. That was checked, not assumed,
    and the check is still run below (see step 5): if a future edition of the workbook introduces
    an "np", the script STOPS rather than silently writing the string into a numeric column.
    This is why the missing-data handling that vis 4 and vis 5 needed is absent here.

5.  THE CLASSIFICATION FILE AND ITS SPELLINGS (changed 2026-10-08). The script used to read
    national_parks_tourism_regions.csv, which was built from Table 16's roster and had no row for
    8 of Table 10's regions (Darwin among them). It now reads
    national_park_tourism_regions_verified.csv - the single source of truth every chart on the
    page now uses (see build_classifications.py) - which covers all 50 Table 10 regions.
    Three regions are spelled differently in the two files; REGION_ALIASES below maps the
    sheet's spelling to the csv's, including the sheet's own typo "Outback Queenland".
    Switching files changed three flags (Peninsula, Central Coast, Lakes: No -> Yes). None of the
    three reaches the top 10 - the closest, Peninsula, is 12th - so vis 6's roster is unchanged.
    Darwin now HAS a row (a capital marked No), so it joins vis 7's comparison group.
"""

from pathlib import Path

# ---- Folder layout (this script lives in scripts/, one level inside the project) ----
ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "raw-data"                 # source files exactly as downloaded
CLASS_DIR = ROOT / "classifications"        # the hand-verified Yes/No region, route and airport csvs
DATA_DIR = ROOT / "docs" / "data"           # processed data the charts load (inside the website)

import csv          # writes the two output files
import io           # lets us force utf-8 on the way out
import os           # used to make sure the data/ folder exists before writing into it

try:
    import openpyxl     # the only way to read .xlsx; not in the standard library
except ImportError:
    raise SystemExit("This script needs openpyxl.  Run:  pip install openpyxl")

# ---- 1. Settings you may want to change -------------------------------------

IN_XLSX = RAW_DIR / "tra-nvs-quarterly-december-2024-2023.xlsx"   # the workbook as downloaded
IN_SHEET = "Table 10"                                   # overnight trips, top 50 regions
IN_FLAGS = CLASS_DIR / "national_park_tourism_regions_verified.csv"  # your verified Yes/No call per region

# Sheet spelling -> csv spelling, for the three regions the two sources name differently (trap 5).
REGION_ALIASES = {
    "Queensland Country": "Southern Queensland Country",
    "Capricorn Coast": "Capricorn",
    "Outback Queenland": "Outback Queensland",     # [sic] - the sheet's own typo
}
OUT_DIR = DATA_DIR                                      # docs/data, the same folder the maps read from
OUT_VIS6 = os.path.join(OUT_DIR, "np_trips_pairs.csv")
OUT_VIS7 = os.path.join(OUT_DIR, "np_exp_per_trip.csv")

# How many national-park regions to plot. Nicholas chose 10: all 36 Yes regions in Table 10 made a
# 36-row bar chart that no longer fit a page and buried the comparison he wanted to make.
# They are the 10 LARGEST by 2024 overnight trips, so "top 10" means "most visited", which is
# the question vis 6 asks. Raise this and both charts grow by the same rows, in the same order.
TOP_N_PARK_REGIONS = 10

# Which measure decides that top 10. Both charts use the SAME roster on purpose: the reader
# learns ten region names once and then sees them again, instead of two unrelated shortlists.
RANK_BY = "trips_2024"

# The comparison group for vis 7. Nicholas chose capital cities.
#
# Two of Australia's capitals are flagged "Yes" in your csv - Canberra and Hobart and the South.
# They are NOT listed here, because a region cannot be in both groups: your csv is the authority
# on what counts as a national-park region, so a capital you flagged Yes belongs in the park
# group. (Canberra lands in the top 10 on its own merits, at 3.10M trips.)
#
# Darwin is INCLUDED. It used to be left out because the old csv had no row for it; the verified
# csv classifies it No (a park-tour gateway, not a park region), which makes it a "No" capital
# like the other five - and vis 10 already counts it as one. Note it is the most extreme value in
# the whole table ($1,877 per trip), which is why VIS7_AXIS_MAX in vis7-lolipop.js was raised.
CAPITAL_COMPARISON_REGIONS = [
    "Sydney",
    "Melbourne",
    "Brisbane",
    "Destination Perth",
    "Adelaide",
    "Darwin",
]

# Group names. These are what the chart's key prints, so they are worded for a reader, not for
# a spreadsheet. Changing them here changes them in the csv and therefore in the chart.
GROUP_PARK = "National park regions"
GROUP_CAPITAL = "Capital cities"

# Optional shorter axis labels. Region names in this sheet are mostly short enough to print in
# full, which is why this is nearly empty - a renamed region is a region a marker cannot look up
# in the source. Add an entry only when a name genuinely will not fit.
SHORT_LABELS = {
    # "Australia's South West": "South West WA",
}

# ---- 2. Where the numbers live in the sheet ---------------------------------
# Row and column positions, 1-based, exactly as Excel counts them. They are written out rather
# than searched for because this sheet's header is split across two merged rows (row 5 names the
# measure, row 6 names the year), so there is no single row of unique column names to match on.
#
#   row 3   title            "10 Overnight trips, visitor nights and regional expenditure..."
#   row 4   "RANKED BY NIGHTS IN REGION FOR YEAR ENDING DECEMBER 2024"      <- trap 2
#   row 5   measure names    Trips ('000) | Visitor nights ('000) | Regional expenditure ($M)...
#   row 6   years            Year Ending December 2023 | Year Ending December 2024 | ...
#   row 7   "Rank"  "Region"
#   rows 8-57   the 50 regions
#   row 59  "Other regions"  <- an aggregate, excluded
#   row 60  "Total"          <- an aggregate, excluded
#   rows 62+    footnotes
FIRST_DATA_ROW = 8
LAST_DATA_ROW = 57          # row 58 is blank, then the two aggregate rows

COL_REGION = 2
COL_TRIPS_2023 = 3
COL_TRIPS_2024 = 4
COL_NIGHTS_2023 = 5
COL_NIGHTS_2024 = 6
COL_EXPENDITURE_2023 = 7     # $ millions
COL_EXPENDITURE_2024 = 8     # $ millions
COL_EXP_PER_TRIP_2023 = 9    # $ per trip, computed by the source
COL_EXP_PER_TRIP_2024 = 10   # $ per trip, computed by the source

# The two aggregate rows, excluded by name as well as by row number so a shifted sheet cannot
# smuggle them in. They are sums of the regions below them and would double-count - the same
# reason prepare_vis45.py drops the "All (Tenured)" rows.
AGGREGATE_ROWS = {"Other regions", "Total"}

# Trips are published in thousands; this turns them into whole trips. See trap 3.
TRIPS_UNIT = 1000

# ---- 3. Read the sheet ------------------------------------------------------

book = openpyxl.load_workbook(IN_XLSX, read_only=True, data_only=True)
if IN_SHEET not in book.sheetnames:
    raise SystemExit("{0} has no sheet called {1!r}".format(IN_XLSX, IN_SHEET))
sheet = book[IN_SHEET]

# openpyxl hands back a tuple per row with 0-based indexing, so every column constant above is
# read as [COL - 1]. Doing the subtraction in one helper keeps that off-by-one in a single place.
def cell(row, column):
    return row[column - 1]


regions = []
not_published = []      # any "np" we meet, collected so the run can fail loudly - see trap 4

for row in sheet.iter_rows(min_row=FIRST_DATA_ROW, max_row=LAST_DATA_ROW, values_only=True):
    name = cell(row, COL_REGION)
    if name is None:
        continue                                    # the blank separator row
    name = str(name).strip()
    if name in AGGREGATE_ROWS:
        continue                                    # see AGGREGATE_ROWS above

    record = {"region": name}
    for key, column in [
        ("trips_2023", COL_TRIPS_2023), ("trips_2024", COL_TRIPS_2024),
        ("nights_2023", COL_NIGHTS_2023), ("nights_2024", COL_NIGHTS_2024),
        ("expenditure_m_2023", COL_EXPENDITURE_2023), ("expenditure_m_2024", COL_EXPENDITURE_2024),
        ("exp_per_trip_2023", COL_EXP_PER_TRIP_2023), ("exp_per_trip_2024", COL_EXP_PER_TRIP_2024),
    ]:
        value = cell(row, column)
        # A suppressed cell arrives as the STRING "np", not as a blank. Recording which cell it
        # was, rather than coercing it to 0 or to None, is what lets step 5 refuse to continue.
        if isinstance(value, str):
            not_published.append("{0} / {1} = {2!r}".format(name, key, value.strip()))
            record[key] = None
        else:
            record[key] = float(value)
    regions.append(record)

print("read {0} regions from {1!r}".format(len(regions), IN_SHEET))

# ---- 4. Read your Yes/No decisions ------------------------------------------
# utf-8-sig, not utf-8: Excel writes a byte-order mark at the start of a csv it saved, and a
# plain utf-8 read turns the first header into "﻿Region", so the "Region" lookup fails and
# every row silently becomes unflagged. (This is the same class of bug as the cp1252 trap in
# prepare_vis45.py - an encoding detail that loses rows without raising anything.)
flags = {}
with io.open(IN_FLAGS, "r", encoding="utf-8-sig", newline="") as handle:
    for row in csv.DictReader(handle):
        region = row["tourism_region"].strip()
        # Anything other than a literal "Yes" counts as not selected. Compared case-insensitively
        # so "yes" in a hand-edited row still registers.
        flags[region] = row["national_park_region"].strip().lower() == "yes"
# Make every sheet spelling findable too (trap 5), so the lookups below need no special cases.
for sheet_name, csv_name in REGION_ALIASES.items():
    if csv_name in flags:
        flags[sheet_name] = flags[csv_name]
# Counted WITHOUT the alias spellings, which are duplicates of real rows.
real = {k: v for k, v in flags.items() if k not in REGION_ALIASES}
print("read {0} Yes/No decisions from {1!r} ({2} marked Yes)".format(
    len(real), IN_FLAGS, sum(1 for v in real.values() if v)))

# ---- 5. The three checks that make this script safe to re-run ---------------
# Each of these would otherwise be a silent wrong answer rather than an error.

# (a) Suppressed cells. Table 10 currently has none; if a future edition does, stop, because
#     every decision below this line assumes complete figures.
if not_published:
    raise SystemExit(
        "This sheet now contains suppressed ('np') cells, which it did not when this script was\n"
        "written. Decide how each chart should show them BEFORE re-running:\n  "
        + "\n  ".join(not_published))

# (b) Regions the flags file does not cover. Not fatal - see trap 5 - but reported every run so
#     the gap never goes unnoticed.
unflagged = [r["region"] for r in regions if r["region"] not in flags]

# (c) Could an unflagged region have made the top 10? If the largest unflagged region is bigger
#     than the smallest region in the chosen top 10, then the shortlist depends on a region you
#     never classified, and the chart cannot be trusted until you do. Checked at the bottom of
#     step 6, once the top 10 is known.

# ---- 6. Choose the rows each chart plots ------------------------------------

# Every region you marked Yes, largest first. sort is on the RANK_BY measure, not on the sheet's
# own rank column, for the reason in trap 2.
park_regions = sorted(
    (r for r in regions if flags.get(r["region"], False)),
    key=lambda r: r[RANK_BY], reverse=True)

top_park = park_regions[:TOP_N_PARK_REGIONS]
print("\ntop {0} national-park regions by {1}:".format(TOP_N_PARK_REGIONS, RANK_BY))
for position, r in enumerate(top_park, 1):
    print("  {0:>2}. {1:<32} {2:>11,.0f} trips".format(
        position, r["region"], r[RANK_BY] * TRIPS_UNIT))
if len(park_regions) > TOP_N_PARK_REGIONS:
    cut = park_regions[TOP_N_PARK_REGIONS]
    print("      (next out: {0}, {1:,.0f} trips)".format(
        cut["region"], cut[RANK_BY] * TRIPS_UNIT))

# Check (c) from step 5, now that the cut-off is known.
if unflagged:
    cutoff = top_park[-1][RANK_BY]
    contenders = [r for r in regions if r["region"] in unflagged and r[RANK_BY] > cutoff]
    print("\n{0} Table 10 regions have no row in {1!r} and were excluded:".format(
        len(unflagged), IN_FLAGS))
    for name in unflagged:
        print("    " + name)
    if contenders:
        raise SystemExit(
            "...and {0} of them is larger than the smallest region in the top {1}, so the\n"
            "shortlist depends on regions you have not classified. Add them to {2!r} first:\n  "
            .format(len(contenders), TOP_N_PARK_REGIONS, IN_FLAGS)
            + "\n  ".join(r["region"] for r in contenders))
    print("    (none of them reaches the top {0} cut-off of {1:,.0f} trips, so the shortlist\n"
          "     above does not depend on classifying them.)".format(
              TOP_N_PARK_REGIONS, cutoff * TRIPS_UNIT))

# The comparison set for vis 7, looked up by name so a typo fails loudly here rather than
# quietly producing a shorter chart.
by_name = {r["region"]: r for r in regions}
comparison = []
for name in CAPITAL_COMPARISON_REGIONS:
    if name not in by_name:
        raise SystemExit(
            "{0!r} is in CAPITAL_COMPARISON_REGIONS but not in {1!r}. Check the spelling against\n"
            "the sheet - it uses 'Destination Perth', not 'Perth'.".format(name, IN_SHEET))
    if flags.get(name, False):
        raise SystemExit(
            "{0!r} is in CAPITAL_COMPARISON_REGIONS but your csv marks it Yes, so it would appear\n"
            "in both groups. Remove it from CAPITAL_COMPARISON_REGIONS, or change the csv."
            .format(name))
    comparison.append(by_name[name])
comparison.sort(key=lambda r: r[RANK_BY], reverse=True)


def short_label(region):
    """The axis label. Falls back to the region's real name, which is usually what you want."""
    return SHORT_LABELS.get(region, region)


# ---- 7. Build vis 6: one row per region, both years side by side ------------
# "Wide" rather than "long" on purpose. The chart needs BOTH years in one row to work out a
# percentage change and to position that figure beyond whichever bar is longer; Vega-Lite can
# fold a wide row into two bars at draw time (and vis6-bar.js does), but it cannot easily go the
# other way without a window transform.
vis6_rows = []
for position, r in enumerate(top_park, 1):
    trips_2023 = r["trips_2023"] * TRIPS_UNIT
    trips_2024 = r["trips_2024"] * TRIPS_UNIT
    change = trips_2024 - trips_2023
    vis6_rows.append({
        "region": r["region"],
        "short_label": short_label(r["region"]),
        "trips_rank_2024": position,          # OUR ranking, not the sheet's - see trap 2
        "trips_2023": int(round(trips_2023)),
        "trips_2024": int(round(trips_2024)),
        "change_trips": int(round(change)),
        # Percentage change, computed here so the chart never has to divide. Guarded against a
        # zero baseline even though no region has one, because an unguarded division is the kind
        # of line that breaks years later on a sheet nobody re-checked.
        "pct_change": round(change / trips_2023 * 100, 1) if trips_2023 else "",
        # Carried through for the tooltip: nights are what makes these OVERNIGHT trips, and a
        # reader checking the chart against the source will look for them.
        "nights_2024": int(round(r["nights_2024"] * TRIPS_UNIT)),
    })

# ---- 8. Build vis 7: one row per plotted region, with its group --------------
# Both groups go in ONE file with a "group" column rather than into two files. The chart sorts
# all of them into a single ranked list and colours by group, which is the whole point: if the
# two groups were separate files they could only be drawn as separate panels, and you could no
# longer see a park region outranking a capital.
vis7_rows = []
for group, members in [(GROUP_PARK, top_park), (GROUP_CAPITAL, comparison)]:
    for r in members:
        trips_2024 = r["trips_2024"] * TRIPS_UNIT
        expenditure_2024 = r["expenditure_m_2024"] * 1e6      # $M -> dollars
        # The sheet publishes expenditure per trip already, and that published figure is what
        # the chart plots - it is the source's own number, so a marker can look it up.
        # This recomputes it anyway, as an arithmetic check on the column mapping in step 2:
        # if COL_EXPENDITURE_2024 or COL_TRIPS_2024 pointed at the wrong column, the two would
        # disagree and the assert below would catch it. (The tolerance is 0.5%, which absorbs
        # the rounding in the published trips and expenditure figures without hiding a real
        # mismatch - a wrong column is out by a factor, not by a fraction of a percent.)
        recomputed = expenditure_2024 / trips_2024
        published = r["exp_per_trip_2024"]
        if abs(recomputed - published) > published * 0.005:
            raise SystemExit(
                "Column mapping looks wrong for {0}: expenditure/trips gives {1:,.2f} but the\n"
                "sheet's own 'Avg. expenditure per trip' column says {2:,.2f}. Re-check the\n"
                "COL_* constants in step 2 against the sheet.".format(
                    r["region"], recomputed, published))

        vis7_rows.append({
            "region": r["region"],
            "short_label": short_label(r["region"]),
            "group": group,
            # Rounded to whole dollars. The source carries 13 decimal places, which are an
            # artefact of dividing two rounded numbers, not precision anyone can use.
            "exp_per_trip_2024": round(published, 0),
            "exp_per_trip_2023": round(r["exp_per_trip_2023"], 0),
            # Both inputs to that average, for the tooltip. A reader who wants to know WHY a
            # region spends more per trip usually wants to know whether it has few expensive
            # trips or many cheap ones, and these two answer that.
            "trips_2024": int(round(trips_2024)),
            "expenditure_m_2024": round(r["expenditure_m_2024"], 1),
            "nights_2024": int(round(r["nights_2024"] * TRIPS_UNIT)),
            # Nights per trip: the usual explanation for a high spend per trip is a longer stay.
            "nights_per_trip_2024": round(r["nights_2024"] / r["trips_2024"], 1),
        })

# ---- 9. Write both files ----------------------------------------------------

os.makedirs(OUT_DIR, exist_ok=True)       # no error if data/ is already there


def write_csv(path, records):
    """Write a list of dicts as a csv, taking the column order from the first record."""
    with io.open(path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(records[0].keys()))
        writer.writeheader()
        writer.writerows(records)
    print("wrote {0}  ({1} rows)".format(path, len(records)))


print()
write_csv(OUT_VIS6, vis6_rows)
write_csv(OUT_VIS7, vis7_rows)

# A short report so you can sanity-check the judgement calls without opening the csv files.
print("\nvis 6 - overnight trips, 2023 -> 2024:")
for r in vis6_rows:
    print("  {0:<32} {1:>11,} -> {2:>11,}   {3:>+6.1f}%".format(
        r["short_label"], r["trips_2023"], r["trips_2024"], r["pct_change"]))

print("\nvis 7 - 2024 expenditure per trip, all plotted regions ranked:")
for r in sorted(vis7_rows, key=lambda x: -x["exp_per_trip_2024"]):
    mark = "*" if r["group"] == GROUP_PARK else " "
    print("  {0} {1:<32} ${2:>7,.0f}   {3:>4.1f} nights/trip".format(
        mark, r["short_label"], r["exp_per_trip_2024"], r["nights_per_trip_2024"]))
print("  (* = {0})".format(GROUP_PARK))
