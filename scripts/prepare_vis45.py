"""
Turn the wide Queensland protected-area sheet into two TIDY csv files.

    indicator-1-1-0-3-1.csv          (input,  1 row per category, 51 year columns)
      |
      +--> data/qld_area_change.csv  (vis 4,  1 row per category, both endpoints side by side)
      +--> data/qld_rank_series.csv  (vis 5,  1 row per category PER YEAR it existed)

Run it with:  python prepare_vis45.py

WHY A PREP SCRIPT AT ALL?
Vega-Lite can reshape data a little (fold, pivot, window...), but this sheet needs work that
is genuinely awkward in a spec: 17 repeating 3-column year blocks, thousands separators inside
quoted fields ("7,176,441"), footnote daggers glued onto category names, and three different
ways of saying "no data". Doing that here keeps the two .js specs about DESIGN, not cleaning -
the same split your prepare_data.py and prepare_map3.py already use for the maps.
"""

from pathlib import Path

# ---- Folder layout (this script lives in scripts/, one level inside the project) ----
ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "raw-data"                 # source files exactly as downloaded
CLASS_DIR = ROOT / "classifications"        # the hand-verified Yes/No region, route and airport csvs
DATA_DIR = ROOT / "docs" / "data"           # processed data the charts load (inside the website)

import csv          # reads the sheet properly, including values quoted because of their commas
import io           # lets us force an encoding when opening the file
import re           # used to strip the footnote daggers off category names
import os           # used to make sure the data/ folder exists before writing into it

# ---- 1. Settings you may want to change -------------------------------------

IN_FILE = RAW_DIR / "indicator-1-1-0-3-1.csv"   # the sheet as downloaded, left untouched
OUT_DIR = DATA_DIR                          # docs/data, the same folder the maps read from
OUT_VIS4 = os.path.join(OUT_DIR, "qld_area_change.csv")
OUT_VIS5 = os.path.join(OUT_DIR, "qld_rank_series.csv")

# The sheet was saved by Excel on a Windows machine, so its non-ASCII characters (the footnote
# daggers after some category names) are cp1252, NOT utf-8. Opening it as utf-8 raises an error
# on those bytes; opening it as cp1252 reads them cleanly and we strip them a few lines below.
IN_ENCODING = "cp1252"

# vis 4 compares two points in time. The sheet names the later years as financial years
# ("2023/24 hectare"), so the END label is written the way the source writes it.
START_YEAR = 2008          # first year in the sheet
END_YEAR = 2024            # last year in the sheet (its hectare column is headed "2023/24")
START_LABEL = "2008"       # what the chart should CALL the start point
END_LABEL = "2023/24"      # what the chart should CALL the end point

# Both totals rows are sums of the rows above them, so plotting them next to their own parts
# would double-count. vis 4's brief excludes them; vis 5 would be dominated by them.
SKIP_ROWS = {"All (Tenured)", "All (Non-tenured)"}

# The sheet's category names are too long to use as axis labels (one is 50 characters), so each
# gets a short display label. EDIT THE RIGHT-HAND SIDE to change what the charts print; the
# left-hand side must keep matching the sheet exactly or the row is dropped as unknown.
SHORT_LABELS = {
    "National park":                                            "National park",
    "National park (scientific)":                               "NP (scientific)",
    "National park (recovery)":                                 "NP (recovery)",
    "National park (Aboriginal land)":                           "NP (Aboriginal land)",
    "National park (Cape York Peninsula Aboriginal land)":        "NP (Cape York Peninsula)",
    "Conservation park / regional park (general)":                "Conservation park",
    "Resources reserve / regional park (resource use area)":      "Resources reserve",
    "Nature refuge":                                            "Nature refuge",
    "Special Wildlife Reserve":                                  "Special Wildlife Reserve",
    "Coordinated conservation area":                              "Coordinated conservation area",
}

# ---- 2. Small helpers -------------------------------------------------------


def clean_name(raw):
    """Strip the footnote markers and stray spaces off a category name.

    'National park (scientific)<dagger>'            -> 'National park (scientific)'
    'Conservation park / regional park (general)^'  -> 'Conservation park / regional park (general)'

    The sheet marks footnotes in TWO ways and both have to go, or SHORT_LABELS stops matching
    and the category is silently dropped from both charts:
      1. non-ASCII daggers, which look like mojibake but are real cp1252 characters; and
      2. ordinary ASCII symbols - '^' on Conservation park, '*' on Resources reserve - which
         survive an "ASCII only" filter precisely BECAUSE they are ASCII.
    """
    # Step 1: drop every non-ASCII character, which removes the dagger-style markers.
    text = re.sub(r"[^\x20-\x7e]", "", raw).strip()
    # Step 2: drop any run of footnote symbols left clinging to the END of the name. Anchoring
    # to the end matters: '/' and '(' inside a name like "Conservation park / regional park
    # (general)" must survive, and only trailing markers are footnotes.
    return re.sub(r"[\^*?#~+]+$", "", text).strip()


def to_number(raw):
    """Turn one cell of the sheet into a float, or None when there is no value.

    Returns None for '' (blank) and for '-'; otherwise removes the thousands separators the
    later years use ("7,176,441") and the stray spaces, then parses what is left.
    None is deliberately DIFFERENT from 0.0 - see the long note in step 4.
    """
    text = (raw or "").strip().replace(",", "").replace(" ", "")
    if text in ("", "-"):              # genuinely empty cell
        return None
    try:
        return float(text)             # handles both '53189' and '7110055.372'
    except ValueError:
        return None                    # anything unparseable is treated as no value


# ---- 3. Read the sheet ------------------------------------------------------

with io.open(IN_FILE, encoding=IN_ENCODING, newline="") as handle:
    rows = list(csv.reader(handle))

header = rows[0]

# The year columns repeat in blocks of three: "<year> number", "<year> hectare area",
# "<year> percent Qld" - but the later years are headed "2020/21 hectare" and "2021/22 percent",
# so the ONLY header we can match reliably is the "number" one. We find each year by its number
# column, then take the hectare column as the one immediately to its right.
# This is why the script never reads the sheet's own percent columns: their headings are
# inconsistent AND their contents are inconsistent (most years hold '0.03%', the 2016 column
# holds 0.0003). Percentage change is recomputed from the hectare figures instead.
year_cols = {}                                   # {2008: (number_index, hectare_index), ...}
for index, head in enumerate(header):
    if head.strip().endswith("number"):          # e.g. '2008 number', '2021 number'
        year = int(head.strip().split()[0])      # the year is the first word of the heading
        year_cols[year] = (index, index + 1)     # hectares always sit one column to the right

years = sorted(year_cols)                        # 2008 ... 2024, in order

# Pull out one tidy record per category: its name, its tenure, and its two value series.
categories = []
for row in rows[1:]:
    name = clean_name(row[0]) if row else ""
    if not name or name in SKIP_ROWS:            # blank spacer rows and the two totals rows
        continue
    if name not in SHORT_LABELS:                 # footnote paragraphs at the bottom of the sheet
        continue
    numbers, hectares = {}, {}
    for year in years:
        n_index, a_index = year_cols[year]
        numbers[year] = to_number(row[n_index] if n_index < len(row) else "")
        hectares[year] = to_number(row[a_index] if a_index < len(row) else "")
    categories.append({
        "category": name,
        "short_label": SHORT_LABELS[name],
        "tenure": row[1].strip(),                # 'Tenured' or 'Non-tenured'
        "numbers": numbers,
        "hectares": hectares,
    })

# ---- 4. THE MISSING-DATA RULE (the one judgement call in this script) -------
#
# The sheet says "this category is not here" in three different ways, and they do NOT mean the
# same thing:
#
#   (a) blank number AND blank hectares
#         National park (scientific) 2014-2015, National park (recovery) 2014-2020.
#         The category was ABOLISHED. It has no parks because it does not exist.
#
#   (b) a literal 0
#         National park (recovery) 2021-2024 (0 parks, 0 ha), National park (Aboriginal land)
#         2021-2022 (0 parks, blank ha). The row is kept in the table for continuity, but
#         zero parks still means the category is not on the ground.
#
#   (c) blank because the category had not been invented yet
#         Special Wildlife Reserve before 2021, National park (Aboriginal land) before 2023.
#
# For vis 5 (ranking by number of parks) all three mean the SAME thing: there is nothing to
# rank, so the category gets no row for that year and its line breaks. That is the brief's
# "break lines where categories are absent or abolished. Do not turn every blank into zero" -
# and note that it cuts the other way too: a literal 0 is not turned into a rank either.
#
# For vis 4 (change in area) (a)/(b) and (c) differ, which is why there is a baseline_note
# column: a category that existed but whose area went to zero really did LOSE that area, while
# a category invented in 2021 started from nothing. Both are honest; conflating them is not.


def exists(number):
    """Did this category exist in this year? Used by vis 5 to decide where lines break."""
    return number is not None and number > 0      # covers (a) blank and (b) literal zero


# ---- 5. Build the vis 4 table (one row per category, two endpoints) ---------

vis4_rows = []
for item in categories:
    hectares, numbers = item["hectares"], item["numbers"]

    start_ha = hectares.get(START_YEAR)
    end_ha = hectares.get(END_YEAR)

    # Which of the three missing-baseline cases is this? The distinction is the whole point of
    # the column: it is what stops an invented-in-2021 category from looking like a 2008 one.
    if start_ha is not None:
        # Normal case: a real 2008 hectare figure.
        start_label, baseline_note = START_LABEL, "reported"
    elif exists(numbers.get(START_YEAR)):
        # The category EXISTED in 2008 (it has a park count) but its area was never published.
        # This is only National park (Cape York Peninsula Aboriginal land): 2 parks, no hectares.
        # Substituting 0 would invent a 2.65 million hectare gain out of a reporting gap, so the
        # baseline moves forward to the first year that actually has an area instead.
        first_year = next((y for y in years if hectares.get(y) is not None), None)
        start_ha = hectares.get(first_year)
        start_label, baseline_note = str(first_year), "area not reported in 2008"
    else:
        # The category did not exist in 2008 at all, so 0 hectares is a FACT, not a guess.
        start_ha, start_label, baseline_note = 0.0, START_LABEL, "category new since 2008"

    # A category abolished before the end year has no end figure; 0 hectares is again a fact.
    if end_ha is None:
        end_ha = 0.0

    change = end_ha - start_ha                                    # the headline number
    # Percentage change is meaningless when the baseline is 0 (you cannot divide by it), so it
    # is left BLANK rather than shown as infinity. The chart prints 'new' for those rows.
    pct = (change / start_ha * 100.0) if start_ha else None

    vis4_rows.append({
        "category": item["category"],
        "short_label": item["short_label"],
        "tenure": item["tenure"],
        "start_year": start_label,
        "start_ha": round(start_ha),
        "end_year": END_LABEL,
        "end_ha": round(end_ha),
        "change_ha": round(change),
        "pct_change": "" if pct is None else round(pct, 1),
        "baseline_note": baseline_note,
    })

# Biggest gain first, biggest loss last. The chart re-sorts on this field anyway, but a file
# that is already in reading order is much easier to check by eye.
vis4_rows.sort(key=lambda r: r["change_ha"], reverse=True)

# ---- 6. Build the vis 5 table (one row per category per year it existed) ----
#
# "segment" is the trick that makes the broken lines work. A Vega-Lite line mark joins its
# points in order, so simply leaving 2014 and 2015 out would draw National park (scientific)
# straight from 2013 to 2016 as if nothing had happened. Numbering each unbroken RUN of years
# and drawing one line per run instead means the gap is a real gap on the page.
vis5_rows = []
for item in categories:
    numbers = item["numbers"]
    segment = 0                 # how many runs of consecutive years we have started so far
    present_years = [y for y in years if exists(numbers[y])]
    if not present_years:
        continue                # category never appears with a park in it; nothing to plot

    # One label per category describing WHAT HAPPENED to it across the whole period. vis 5
    # colours its lines by this, so the colour carries the classification story the footnotes
    # of the sheet describe, instead of just repeating the category name the labels already give.
    # The order of these tests matters: the first one that fits wins.
    if present_years[-1] < END_YEAR:
        # Stops before the last year and never comes back: National park (recovery), gone in 2014.
        status = "Abolished"
    elif present_years[0] > START_YEAR:
        # Appears for the first time partway through: Special Wildlife Reserve (2021),
        # National park (Aboriginal land) (2023).
        status = "New since 2008"
    elif len(present_years) != present_years[-1] - present_years[0] + 1:
        # Present at both ends but with a hole in between, which only happens when a category is
        # abolished and later brought back: National park (scientific), gone 2014-15, back 2016.
        status = "Abolished and reinstated"
    else:
        status = "Continuous since 2008"

    for year in years:
        if not exists(numbers[year]):
            continue            # absent or abolished: no row at all, so no rank and no line
        # A run starts whenever the PREVIOUS year was absent (or this is the first year).
        if year == present_years[0] or not exists(numbers.get(year - 1)):
            segment += 1
        vis5_rows.append({
            "category": item["category"],
            "short_label": item["short_label"],
            "tenure": item["tenure"],
            "status": status,
            "year": year,
            "number": int(numbers[year]),
            # Unique per run per category, so one category's two runs never join up.
            "segment": "{0} #{1}".format(item["short_label"], segment),
            # Flags the chart uses to put a text label at each end of a category's life.
            "is_first": 1 if year == present_years[0] else 0,
            "is_last": 1 if year == present_years[-1] else 0,
        })

# ---- 7. Write both files ----------------------------------------------------

os.makedirs(OUT_DIR, exist_ok=True)       # no error if data/ is already there


def write_csv(path, records):
    """Write a list of dicts as a csv, taking the column order from the first record."""
    with io.open(path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(records[0].keys()))
        writer.writeheader()
        writer.writerows(records)
    print("wrote {0}  ({1} rows)".format(path, len(records)))


write_csv(OUT_VIS4, vis4_rows)
write_csv(OUT_VIS5, vis5_rows)

# A short report so you can sanity-check the judgement calls without opening the csv files.
print("\nvis 4 - change in area, 2008 to 2023/24:")
for r in vis4_rows:
    pct = "{0:+.1f}%".format(r["pct_change"]) if r["pct_change"] != "" else "   new  "
    flag = "" if r["baseline_note"] == "reported" else "   <- " + r["baseline_note"]
    print("  {0:<30} {1:>10,} -> {2:>10,}  {3:>+11,} ha  {4}{5}".format(
        r["short_label"], r["start_ha"], r["end_ha"], r["change_ha"], pct, flag))

print("\nvis 5 - how many categories are ranked each year:")
for year in years:
    present = [r["short_label"] for r in vis5_rows if r["year"] == year]
    print("  {0}: {1} categories".format(year, len(present)))
