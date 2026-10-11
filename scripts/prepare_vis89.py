"""
Turn the two BITRE aviation workbooks into two TIDY csv files for vis 8 and vis 9.

    TopRoutesJuly2014June2026_0.xlsx, sheet "Top Routes"          (monthly traffic per city pair)
    national_park_routes.csv                                      (your Yes/No call per route)
      |
      +--> data/np_route_monthly.csv       (vis 8, 1 row per route per month, 5 routes x 144 months)

    Aviation-Airport Traffic Data-May 2026.xlsx, sheet "Airport Passengers"
    national_park_airport_classification.csv                      (your Yes/No call per airport)
      |
      +--> data/np_gateway_seasonality.csv (vis 9, 1 row per gateway per month of 2024)

Run it with:  python prepare_vis89.py
Needs openpyxl (pip install openpyxl) because both sources are .xlsx.

WHY A PREP SCRIPT AT ALL?
Same reason as prepare_vis45.py and prepare_vis67.py: Vega-Lite cannot read .xlsx, and both
sheets need cleaning that is awkward to express in a spec (see TRAPS). Doing it here keeps the
two .js files about colour, layout and sight lines.

--------------------------------------------------------------------------------------------
TRAPS - read before changing anything
--------------------------------------------------------------------------------------------
1.  THE "TOP ROUTES" PASSENGER COLUMN IS NOT ALWAYS A NUMBER. 1,871 cells hold TEXT instead:
      "Data not available for release."   1,781 cells - the source SUPPRESSED the figure
      "COVID19 No flights"                     90 cells - the source says NOTHING FLEW
    These mean opposite things, and your instruction for vis 8 separates them exactly:
    "a missing route-month should remain missing unless the source establishes that it
    represents zero traffic". The COVID label IS the source establishing zero, so it becomes 0.
    The suppression label is NOT - a suppressed month could have had thousands of passengers -
    so it stays blank, and the chart draws a GAP there instead of a dive to zero.
    A plain float() on the column would crash on the first text cell; a "treat text as 0" shortcut
    would invent a collapse in traffic that never happened. Both are avoided below.

2.  ONE DUPLICATED ROW. Adelaide-Hobart, November 2025 appears twice in the sheet (both copies
    say "Data not available for release."). It is not one of the plotted routes, but the script
    de-duplicates identical rows anyway, and STOPS if two copies of the same route-month ever
    disagree, because then it cannot know which one is right.

3.  THE ROUTE LIST IN YOUR CSV WAS CHECKED AGAINST THE SHEET, NOT TRUSTED. All 80 route codes
    in national_park_routes.csv are the 80 in the sheet (none missing, none extra), and every
    "source_row_count" matches the number of months the sheet holds for that route.
    The csv's one stated ASSUMPTION - that "QNA" is Ballina - is CONFIRMED by the sheet itself,
    whose own route-name column prints "Ballina - Sydney" and "Ballina - Melbourne". The script
    re-checks the route list every run.

4.  SEVERAL SELECTED ROUTES HAVE NO 2024 FIGURE AT ALL. Fourteen "Yes" routes (e.g. Sydney -
    Townsville, Ayers Rock - Sydney) are suppressed for all twelve months of 2024, so they cannot
    be ranked. That does NOT change the top five: the most any of them has ever published in a
    calendar year is about 196,000 passengers (Ayers Rock - Sydney, 2019), against a fifth-place
    cut-off of about 1,009,000. The script re-checks that claim every run and stops if it fails.

5.  BOTH CLASSIFICATION CSVs ARE GENERATED - DO NOT EDIT THEM. national_park_routes.csv and
    national_park_airport_classification.csv are rebuilt by build_classifications.py from
    national_park_tourism_regions_verified.csv, so that every chart uses ONE definition of a
    national-park region. To change a classification, edit the verified csv and re-run
    build_classifications.py, then this script.
    (The airport file used to have one row per sheet row with a "source_row" column that no
    longer lined up with the sheet; it now has one row per airport. This script joins on the
    airport NAME and stops if any airport ever carries two different flags.)

6.  AIRPORT PASSENGERS ARE MOVEMENTS, NOT PEOPLE. The sheet's Notes tab: "Each domestic
    passenger generates two passenger movements (a departure and an arrival)." One return trip
    to Cairns from Sydney is two movements at Cairns. That is why vis 9's labels must say
    "passenger movements" rather than "visitors" or "passengers".
"""

import calendar      # month names and days-in-month, so no lookup table has to be typed out
from pathlib import Path

# ---- Folder layout (this script lives in scripts/, one level inside the project) ----
ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "raw-data"                 # source files exactly as downloaded
CLASS_DIR = ROOT / "classifications"        # the hand-verified Yes/No region, route and airport csvs
DATA_DIR = ROOT / "docs" / "data"           # processed data the charts load (inside the website)

import collections   # Counter and defaultdict, for the checks
import csv           # reads your two classification files and writes the two outputs
import io            # lets us force the encoding on the way in and out
import os            # makes sure the data/ folder exists before writing into it

try:
    import openpyxl  # the only way to read .xlsx; not in the standard library
except ImportError:
    raise SystemExit("This script needs openpyxl.  Run:  pip install openpyxl")

# ---- 1. Settings you may want to change -------------------------------------

ROUTES_XLSX = RAW_DIR / "TopRoutesJuly2014June2026_0.xlsx"
ROUTES_SHEET = "Top Routes"
ROUTES_FLAGS = CLASS_DIR / "national_park_routes.csv"

AIRPORTS_XLSX = RAW_DIR / "Aviation-Airport Traffic Data-May 2026.xlsx"
AIRPORTS_SHEET = "Airport Passengers"
AIRPORTS_FLAGS = CLASS_DIR / "national_park_airport_classification.csv"

OUT_DIR = DATA_DIR
OUT_VIS8 = os.path.join(OUT_DIR, "np_route_monthly.csv")
OUT_VIS9 = os.path.join(OUT_DIR, "np_gateway_seasonality.csv")

# Vis 8: how many routes, and which calendar year chooses them. Your stub comment fixes both:
# "the top five routes ... using 2024 passenger totals, and keep that selection fixed".
TOP_N_ROUTES = 5
RANK_YEAR = 2024

# Vis 9: which year is shown. "each month's share of that airport's annual ... traffic in 2024".
HEATMAP_YEAR = 2024

# Which flag value counts as "selected" in both of your csv files. "Needs review" and
# "Not applicable" are deliberately NOT selected - the same rule as vis 6/7, where a region your
# csv does not mark Yes is left out rather than guessed at.
SELECTED = "Yes"

# Airport latitudes, in decimal degrees (negative = south). Used ONLY to order the heatmap's rows
# from north to south, so the reader can see whether the timing of the peak follows climate.
# Approximate airport-reference-point positions, rounded to 0.1 degree - far finer than the
# ordering needs. If you add an airport, add its latitude here or the script stops.
# (vis9-heatmap.js has a constant to order rows by peak month or by size instead.)
GATEWAY_LATITUDE = {
    "CAIRNS": -16.9,
    "TOWNSVILLE": -19.3,
    "HAMILTON ISLAND": -20.4,
    "PROSERPINE": -20.5,
    "KARRATHA": -20.7,      # in WA - the order is by latitude only, whatever the state
    "MACKAY": -21.2,
    "ROCKHAMPTON": -23.4,
    "SUNSHINE COAST": -26.6,
    "GOLD COAST": -28.2,
    "BALLINA": -28.8,
    "CANBERRA": -35.3,
    "LAUNCESTON": -41.5,
    "HOBART": -42.8,
}

# ---- 2. Where the numbers live in each sheet --------------------------------
# 1-based column numbers, as Excel counts them. Written out rather than searched for because both
# sheets have multi-row merged headers (see the row maps below), so there is no single header row
# of unique names to match on.
#
# "Top Routes":
#   rows 1-13   title, notes, and a two-row header split across merged cells
#   row 14      "City Pair Route" | | "Year" | "Month" | "Revenue" | "Aircraft" | ...
#   row 15      "Industry Domestic/Regional" | | | | "Passengers" | "Trips (a)" | ...
#   rows 16+    one row per route per month, July 2014 - June 2026
# Columns M-S repeat the measures as MOVING ANNUAL TOTALS - not used, because a 12-month rolling
# sum would smear every month across the year after it and hide exactly the monthly shape vis 8
# is meant to show.
R_FIRST_DATA_ROW = 16
R_COL_PORT_1 = 1          # e.g. "OOL"  (the pair is always written alphabetically)
R_COL_PORT_2 = 2          # e.g. "SYD"
R_COL_YEAR = 3
R_COL_MONTH = 4
R_COL_PASSENGERS = 5      # "Revenue Passengers" - both directions combined (sheet note, row 3)
R_COL_ROUTE_NAME = 21     # e.g. "Gold Coast - Sydney" - the source's own label, used on the chart

# The two text values that can appear in the passenger column - see trap 1. Matched after
# stripping spaces and in lower case, so a trailing space in a future edition does not matter.
SUPPRESSED_TEXT = "data not available for release."
NO_FLIGHTS_TEXT = "covid19 no flights"

# Status words written into the csv. These are what the tooltip prints, so they are worded for a
# reader rather than for a spreadsheet.
STATUS_PUBLISHED = "Published"
STATUS_NO_FLIGHTS = "No flights (COVID-19), shown as 0"
STATUS_SUPPRESSED = "Not released by source, left blank"
STATUS_ABSENT = "Not in source, left blank"

# "Airport Passengers":
#   row 5   DOMESTIC and REGIONAL AIRLINES | INTERNATIONAL AIRLINES | TOTAL PASSENGERS
#   row 7   AIRPORT | Year | Month | INBOUND | OUTBOUND | TOTAL | (repeated per sector)
#   rows 8+ one row per airport per month, Jan 2009 - May 2026
A_FIRST_DATA_ROW = 8
A_COL_AIRPORT = 1
A_COL_YEAR = 2
A_COL_MONTH = 3
A_COL_DOMESTIC_TOTAL = 6  # DOMESTIC and REGIONAL, inbound + outbound - the measure you asked for


def cell(row, column):
    """openpyxl rows are 0-based tuples; every COL constant above is 1-based like Excel."""
    return row[column - 1]


def read_flags(path, key_column, flag_column):
    """Read one of your classification csv files into {key: flag}.

    utf-8-sig rather than utf-8 because Excel puts a byte-order mark at the start of csv files it
    saves, which would otherwise glue itself to the first header name and make every lookup fail.
    Stops if the same key is given two different flags - see trap 5.
    """
    seen = collections.defaultdict(set)
    with io.open(path, "r", encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle):
            seen[row[key_column].strip()].add(row[flag_column].strip())
    clashes = {k: v for k, v in seen.items() if len(v) > 1}
    if clashes:
        raise SystemExit("{0}: these keys have more than one flag: {1}".format(path, clashes))
    return {k: v.pop() for k, v in seen.items()}


# ============================================================================
# VIS 8 - monthly passengers on the top five selected routes
# ============================================================================

# ---- 3. Read the route sheet ------------------------------------------------
book = openpyxl.load_workbook(ROUTES_XLSX, read_only=True, data_only=True)
sheet = book[ROUTES_SHEET]

route_name = {}                         # "OOL-SYD" -> "Gold Coast - Sydney"
cells = {}                              # ("OOL-SYD", 2024, 3) -> number, or one of the two texts
for row in sheet.iter_rows(min_row=R_FIRST_DATA_ROW, values_only=True):
    year = cell(row, R_COL_YEAR)
    # A data row always has two airport codes and an integer year. Anything else (a stray note
    # or a blank line) is skipped rather than parsed.
    if not (cell(row, R_COL_PORT_1) and cell(row, R_COL_PORT_2) and isinstance(year, int)):
        continue
    pair = "{0}-{1}".format(cell(row, R_COL_PORT_1).strip(), cell(row, R_COL_PORT_2).strip())
    key = (pair, year, int(cell(row, R_COL_MONTH)))
    value = cell(row, R_COL_PASSENGERS)
    if isinstance(value, str):
        value = value.strip().lower()   # normalise so the comparisons in step 5 are exact
    # Trap 2: an identical duplicate is harmless and ignored; a conflicting one is fatal.
    if key in cells and cells[key] != value:
        raise SystemExit("Two different values for {0}: {1!r} and {2!r}".format(
            key, cells[key], value))
    cells[key] = value
    route_name[pair] = str(cell(row, R_COL_ROUTE_NAME)).strip()

months_in_sheet = sorted({(y, m) for (_, y, m) in cells})
print("read {0} route-months for {1} routes, {2}-{3:02d} to {4}-{5:02d}".format(
    len(cells), len(route_name), *months_in_sheet[0], *months_in_sheet[-1]))

# Any text other than the two known labels means the source has changed its conventions, and
# the right handling for it has to be decided by a person, not guessed here.
unknown = {v for v in cells.values() if isinstance(v, str)} - {SUPPRESSED_TEXT, NO_FLIGHTS_TEXT}
if unknown:
    raise SystemExit("Unrecognised text in the passenger column: {0}".format(unknown))

# ---- 4. Read your route classification and check it against the sheet --------
route_flags = read_flags(ROUTES_FLAGS, "route_pair", "national_park_route")
if set(route_flags) != set(route_name):
    raise SystemExit("Route list mismatch.\n  in sheet only: {0}\n  in csv only:   {1}".format(
        sorted(set(route_name) - set(route_flags)), sorted(set(route_flags) - set(route_name))))
print("read {0} route flags from {1!r} ({2} marked {3!r}) - route list matches the sheet".format(
    len(route_flags), ROUTES_FLAGS, sum(v == SELECTED for v in route_flags.values()), SELECTED))


def is_number(value):
    """True for a real passenger count. bool is excluded because Python treats True as 1."""
    return isinstance(value, (int, float)) and not isinstance(value, bool)


# ---- 5. Rank the selected routes by their 2024 total --------------------------
# A route only gets a 2024 total if ALL twelve months are published numbers. Summing whatever
# months happen to be published would quietly rank a route with three suppressed months below
# one with none, purely because of the suppression.
annual = {}
unrankable = []
for pair, flag in route_flags.items():
    if flag != SELECTED:
        continue
    values = [cells.get((pair, RANK_YEAR, m)) for m in range(1, 13)]
    if all(is_number(v) for v in values):
        annual[pair] = sum(values)
    else:
        unrankable.append(pair)

ranked = sorted(annual, key=annual.get, reverse=True)
top_routes = ranked[:TOP_N_ROUTES]
cutoff = annual[top_routes[-1]]

print("\ntop {0} selected routes by {1} passengers:".format(TOP_N_ROUTES, RANK_YEAR))
for position, pair in enumerate(top_routes, 1):
    print("  {0}. {1:<8} {2:<28} {3:>11,}".format(position, pair, route_name[pair], annual[pair]))
print("      (next out: {0}, {1:,})".format(ranked[TOP_N_ROUTES], annual[ranked[TOP_N_ROUTES]]))

# Trap 4: could a route with no complete 2024 have beaten the cut-off? Its best ever published
# CALENDAR-YEAR total is the most generous evidence the sheet offers. If even that is below the
# cut-off, leaving it out cannot have changed the top five.
for pair in unrankable:
    yearly = collections.Counter()
    for (p, y, _), v in cells.items():
        if p == pair and is_number(v):
            yearly[y] += v
    best = max(yearly.values()) if yearly else 0
    if best >= cutoff:
        raise SystemExit(
            "{0} has no complete {1} but once carried {2:,} passengers in a year, above the\n"
            "cut-off of {3:,}. Decide how to treat it before re-running.".format(
                pair, RANK_YEAR, best, cutoff))
print("  {0} selected routes have no complete {1}; none has ever published a year above the\n"
      "  cut-off, so leaving them out cannot change the top {2}.".format(
          len(unrankable), RANK_YEAR, TOP_N_ROUTES))

# ---- 6. Build vis 8's rows: every route x every month, gaps written as gaps ---
# One row for EVERY month in the sheet's range, for every chosen route, even where the sheet has
# no number. That is what lets the chart draw a BREAK: Vega-Lite can only break a line at a row
# whose value is blank. If the row were simply missing, it would join the months either side of
# it with a straight line and invent the traffic in between.
vis8_rows = []
for rank, pair in enumerate(top_routes, 1):
    # "Gold Coast - Sydney" -> "Gold Coast – Sydney". An en dash is the correct mark for a span
    # between two places; the sheet uses a hyphen only because it is plain text.
    label = route_name[pair].replace(" - ", " – ")
    for (year, month) in months_in_sheet:
        value = cells.get((pair, year, month))
        if is_number(value):
            passengers, status = int(value), STATUS_PUBLISHED
        elif value == NO_FLIGHTS_TEXT:
            passengers, status = 0, STATUS_NO_FLIGHTS        # the source establishes zero
        elif value == SUPPRESSED_TEXT:
            passengers, status = "", STATUS_SUPPRESSED       # unknown, NOT zero
        else:
            passengers, status = "", STATUS_ABSENT           # route not listed that month
        vis8_rows.append({
            "route_pair": pair,
            "route_label": label,
            "rank_2024": rank,
            "pax_2024": annual[pair],
            # ISO date on the first of the month: Vega-Lite parses this as a real date, which is
            # what a temporal x axis needs.
            "month": "{0}-{1:02d}-01".format(year, month),
            "passengers": passengers,
            "status": status,
        })

# The window of months the source labels "COVID19 No flights" on ANY route - printed so that the
# shaded band in vis8-line.js can be checked against the source rather than against memory.
covid_months = sorted((y, m) for (_, y, m), v in cells.items() if v == NO_FLIGHTS_TEXT)
print("\n'COVID19 No flights' appears from {0}-{1:02d} to {2}-{3:02d} across all routes".format(
    *covid_months[0], *covid_months[-1]))
status_count = collections.Counter(r["status"] for r in vis8_rows)
print("vis 8 status of each plotted route-month: {0}".format(dict(status_count)))

# ============================================================================
# VIS 9 - month-by-month share of each gateway's 2024 domestic traffic
# ============================================================================

# ---- 7. Read the airport sheet ------------------------------------------------
book = openpyxl.load_workbook(AIRPORTS_XLSX, read_only=True, data_only=True)
sheet = book[AIRPORTS_SHEET]

movements = {}                          # ("CAIRNS", 3) -> domestic passenger movements, 2024
for row in sheet.iter_rows(min_row=A_FIRST_DATA_ROW, values_only=True):
    airport, year = cell(row, A_COL_AIRPORT), cell(row, A_COL_YEAR)
    if not airport or year != HEATMAP_YEAR:
        continue
    movements[(str(airport).strip(), int(cell(row, A_COL_MONTH)))] = cell(row, A_COL_DOMESTIC_TOTAL)

airport_flags = read_flags(AIRPORTS_FLAGS, "airport", "national_park_tourism")
gateways = sorted(a for a, f in airport_flags.items() if f == SELECTED)
print("\nread {0} airport flags from {1!r}; selected gateways: {2}".format(
    len(airport_flags), AIRPORTS_FLAGS, ", ".join(gateways)))

# ---- 8. Check every selected gateway has a complete, numeric 2024 -------------
# It does today (all 13 x 12 cells are whole numbers, none blank, none text). A share of an
# annual total is meaningless if a month is missing - the other eleven would all be inflated -
# so if a future edition suppresses one, stop and decide rather than divide by a short year.
problems = []
for airport in gateways:
    if airport not in GATEWAY_LATITUDE:
        problems.append("{0}: no latitude in GATEWAY_LATITUDE".format(airport))
    for month in range(1, 13):
        if not is_number(movements.get((airport, month))):
            problems.append("{0} {1}-{2:02d}: {3!r}".format(
                airport, HEATMAP_YEAR, month, movements.get((airport, month))))
if problems:
    raise SystemExit("Cannot build vis 9:\n  " + "\n  ".join(problems))

# ---- 9. Build vis 9's rows ------------------------------------------------------
# Long format: one row per gateway per month. The share is computed HERE, within each airport,
# which is the normalisation your stub asks for: every row of the heatmap then sums to 100%, so
# Gold Coast's 5.6 million movements cannot drown out Ballina's 0.6 million in the colour scale.
vis9_rows = []
for airport in gateways:
    monthly = [movements[(airport, m)] for m in range(1, 13)]
    total = sum(monthly)
    peak_month = monthly.index(max(monthly)) + 1            # first month if two tie
    for month, value in enumerate(monthly, 1):
        vis9_rows.append({
            "airport": airport.title(),                     # "SUNSHINE COAST" -> "Sunshine Coast"
            "latitude": GATEWAY_LATITUDE[airport],
            "month_num": month,
            "month": calendar.month_abbr[month],            # "Jan" ... "Dec"
            "days_in_month": calendar.monthrange(HEATMAP_YEAR, month)[1],   # 2024 is a leap year
            "movements": int(value),
            "annual_movements": int(total),
            # Kept as a FRACTION (0.108, not 10.8) so Vega-Lite's "%" format can print it.
            "share": round(value / total, 5),
            "peak_month_num": peak_month,
            # 1 or 0, NOT Python's True/False. The csv would hold the words "True" and "False",
            # and Vega's boolean parser treats any non-empty string other than lower-case
            # "false" as TRUE - so "False" would read as true and every cell would be marked as
            # a peak. A number cannot be misread.
            "is_peak": 1 if month == peak_month else 0,
        })

# ---- 10. Write both files --------------------------------------------------------
os.makedirs(OUT_DIR, exist_ok=True)


def write_csv(path, records):
    """Write a list of dicts as a csv, taking the column order from the first record."""
    with io.open(path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(records[0].keys()))
        writer.writeheader()
        writer.writerows(records)
    print("wrote {0}  ({1} rows)".format(path, len(records)))


print()
write_csv(OUT_VIS8, vis8_rows)
write_csv(OUT_VIS9, vis9_rows)

# A short report so the shares can be sanity-checked without opening the csv.
print("\nvis 9 - share of {0} domestic passenger movements by month (%):".format(HEATMAP_YEAR))
print("  {0:<15}".format("") + "".join("{0:>6}".format(calendar.month_abbr[m]) for m in range(1, 13)))
for airport in sorted(gateways, key=lambda a: -GATEWAY_LATITUDE[a]):
    shares = [r["share"] for r in vis9_rows if r["airport"] == airport.title()]
    print("  {0:<15}".format(airport.title()) + "".join("{0:>6.1f}".format(s * 100) for s in shares))
print("  (an even spread would be {0:.1f}% a month)".format(100 / 12))
