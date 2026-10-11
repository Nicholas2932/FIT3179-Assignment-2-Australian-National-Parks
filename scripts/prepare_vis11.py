"""
Turn Table 4 of the National Visitor Survey workbook into the geometry of a sankey diagram, for
vis 11 (where do interstate overnight trips to each state come from?).

    tra-nvs-quarterly-december-2024-2023.xlsx, sheet "Table 4", year ending December 2024
      |
      +--> data/np_trip_flows_nodes.csv   (vis 11, 1 row per node: the bars on each side)
      +--> data/np_trip_flows_links.csv   (vis 11, 1 row per link per sample point along its curve)

Run it with:  python prepare_vis11.py
Needs openpyxl (pip install openpyxl) because the source is .xlsx.

WHY THE GEOMETRY IS COMPUTED HERE
Vega-Lite has no sankey mark. A sankey is bars (the nodes) joined by ribbons (the links), and
every ribbon's position depends on every other ribbon at the same node - a cumulative layout that
is clumsy to express as Vega-Lite transforms. So this script lays the diagram out in TRIP UNITS
(thousands of trips on the y axis), and the chart only has to draw it. Because every position is
in trips, the chart's ONE linear y scale turns trips into pixels everywhere: a link 1 million
trips wide is the same thickness wherever it is, which is your "same scale throughout".

DECISIONS (Nicholas, 2026-10-08)
  • Year ending December 2024 (Table 4 also holds 2023), to match vis 6/7.
  • Interstate trips only: the diagonal (a state's own residents) is left out.
  • All eight destinations.
  • "np" cells are never drawn as zero and never guessed one by one. For each destination, the
    source's published TOTAL minus its published cells is drawn as ONE grey link from a
    "not published" node - see trap 3.

--------------------------------------------------------------------------------------------
TRAPS IN THIS TABLE - read before changing anything
--------------------------------------------------------------------------------------------
1.  ROWS ARE DESTINATIONS, COLUMNS ARE ORIGINS. Footnote d: "The numbers in the rows represent
    the state or territory where a visitor went. The numbers across the columns represent the
    state or territory of origin." Reading it the other way round reverses every arrow in the
    diagram and still produces a perfectly plausible-looking chart.

2.  TRIPS ARE IN THOUSANDS ('000). Multiplied by 1,000 on the way out so the csv holds whole trips.

3.  "np" = NOT PUBLISHED ("survey error too high"). 15 interstate cells in 2024. Each
    destination row also has a published total (the "AUST" column), which counts ALL its
    visitors - including the np ones, and (footnote b) visitors whose origin "could not be
    allocated to a state". So for a destination:
        total - intrastate - published interstate cells  =  np cells + unallocated
    That remainder is drawn as one grey link. It is a SUBTRACTION of published figures, not an
    estimate: no individual np cell is given a value. A destination with no np cells (QLD) gets
    no grey link - its remainder is only rounding.

4.  THE ACT CANNOT HAVE A GREY LINK. Its intrastate cell (Canberrans travelling within the ACT)
    is itself np, so its remainder mixes intrastate and interstate trips and cannot be split.
    Drawing it would break the interstate-only rule, so the ACT shows its three published links
    and nothing else. The notes under the chart say so.

5.  AN ORIGIN BAR IS ONLY ITS PUBLISHED LINKS. NSW residents' trips to the NT are np, so NSW's
    bar is the sum of its published interstate links and is slightly SHORT of the truth. That is
    unavoidable without guessing, and the tooltip calls it "published".
"""

from pathlib import Path

# ---- Folder layout (this script lives in scripts/, one level inside the project) ----
ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "raw-data"                 # source files exactly as downloaded
CLASS_DIR = ROOT / "classifications"        # the hand-verified Yes/No region, route and airport csvs
DATA_DIR = ROOT / "docs" / "data"           # processed data the charts load (inside the website)

import csv          # writes the two output files
import io           # forces utf-8 on the way out
import os           # makes sure data/ exists

try:
    import openpyxl  # the only way to read .xlsx; not in the standard library
except ImportError:
    raise SystemExit("This script needs openpyxl.  Run:  pip install openpyxl")

# ---- 1. Settings you may want to change -------------------------------------
IN_XLSX = RAW_DIR / "tra-nvs-quarterly-december-2024-2023.xlsx"
IN_SHEET = "Table 4"
YEAR_HEADING = "YEAR ENDING DECEMBER 2024"       # the block of the sheet to read
MEASURE_HEADING = "Trips to state/territory"     # the first sub-table under that heading
OUT_NODES = os.path.join(DATA_DIR, "np_trip_flows_nodes.csv")
OUT_LINKS = os.path.join(DATA_DIR, "np_trip_flows_links.csv")
TRIPS_UNIT = 1000                                # trap 2

# The eight states and territories, in the SAME order the sheet uses for its columns and rows
# (the ABS's standard order). Both sides of the diagram use this order, top to bottom, so a reader
# finds a state at the same height on either side. Each entry: (code, full name as in the rows).
STATES = [
    ("NSW", "New South Wales"), ("VIC", "Victoria"), ("QLD", "Queensland"),
    ("SA", "South Australia"), ("WA", "Western Australia"), ("TAS", "Tasmania"),
    ("NT", "Northern Territory"), ("ACT", "Australian Capital Territory"),
]
# Column of each origin's figure (1-based, as Excel counts): B = NSW ... I = ACT, J = AUST total.
FIRST_ORIGIN_COLUMN = 2
TOTAL_COLUMN = 10

# Destinations whose np remainder must NOT be drawn - see trap 4.
NO_REMAINDER = {"ACT"}

# The extra node on the left for the grey links. Its code is what the chart tests for.
NP_CODE = "NP"
NP_NAME = "Not published or unallocated"

# Empty space between neighbouring nodes, in THOUSANDS of trips (the layout's own unit). 900 is
# about 2.6% of the diagram's height - enough that the smallest nodes (NT as an origin is only
# 218,000 trips, a sliver) keep a clear gap and room for a one-line label.
NODE_GAP = 900

# How many points each link's curve is sampled at. The curve is drawn as straight segments
# between them; 40 is smooth at any width this page uses.
CURVE_SAMPLES = 40

# ---- 2. Read the 2024 trips block ------------------------------------------------
rows = list(openpyxl.load_workbook(IN_XLSX, read_only=True, data_only=True)[IN_SHEET]
            .iter_rows(values_only=True))
year_row = next(i for i, r in enumerate(rows) if r[0] == YEAR_HEADING)
measure_row = next(i for i in range(year_row, len(rows))
                   if isinstance(rows[i][0], str) and rows[i][0].startswith(MEASURE_HEADING))

cells = {}       # (origin code, destination code) -> thousands of trips, or None for np
totals = {}      # destination code -> its published total ("AUST" column)
name_to_code = {name: code for code, name in STATES}
for r in rows[measure_row + 1: measure_row + 1 + len(STATES)]:
    destination = name_to_code.get(str(r[0]).strip())       # trap 1: the ROW is the destination
    if destination is None:
        raise SystemExit("Unexpected row in {0}: {1!r}".format(IN_SHEET, r[0]))
    totals[destination] = float(r[TOTAL_COLUMN - 1])
    for j, (origin, _) in enumerate(STATES):
        value = r[FIRST_ORIGIN_COLUMN - 1 + j]
        if isinstance(value, str) and value.strip().lower() == "np":
            cells[(origin, destination)] = None              # kept as NOT PUBLISHED, never 0
        elif isinstance(value, (int, float)):
            cells[(origin, destination)] = float(value)
        else:
            raise SystemExit("Unexpected cell {0}->{1}: {2!r}".format(origin, destination, value))

# ---- 3. The links: published interstate cells, plus one np remainder per destination ----
codes = [c for c, _ in STATES]
links = []          # dicts: origin, destination, trips (thousands), np_origins
np_report = []
np_lists = {}       # destination -> its np origins, kept for the node tooltips
for d in codes:
    published = [(o, cells[(o, d)]) for o in codes if o != d and cells[(o, d)] is not None]
    np_origins = [o for o in codes if o != d and cells[(o, d)] is None]
    np_lists[d] = np_origins
    for o, v in published:
        links.append({"origin": o, "destination": d, "trips": v, "np_origins": ""})
    if np_origins and d not in NO_REMAINDER:
        intrastate = cells[(d, d)]
        remainder = totals[d] - intrastate - sum(v for _, v in published)
        if remainder < 0:
            raise SystemExit("Negative np remainder for {0}: {1:.1f}".format(d, remainder))
        links.append({"origin": NP_CODE, "destination": d, "trips": remainder,
                      "np_origins": ", ".join(np_origins)})
        np_report.append((d, remainder, np_origins))
    elif np_origins:
        np_report.append((d, None, np_origins))

# ---- 4. Lay out the nodes: stacked top to bottom, with gaps -------------------------
# A node's height is the sum of the links touching it on that side.
left_order = codes + [NP_CODE]               # the np node goes LAST, at the bottom: it is not a place
right_order = codes
left_size = {c: sum(l["trips"] for l in links if l["origin"] == c) for c in left_order}
right_size = {c: sum(l["trips"] for l in links if l["destination"] == c) for c in right_order}
left_order = [c for c in left_order if left_size[c] > 0]     # drop a node with no links at all


def stack(order, size):
    """Top and bottom of each node, in thousands of trips from the top of the diagram."""
    y, out = 0.0, {}
    for c in order:
        out[c] = (y, y + size[c])
        y += size[c] + NODE_GAP
    return out


left_pos = stack(left_order, left_size)
right_pos = stack(right_order, right_size)

# Within a node, its links are stacked in the order of the node at their OTHER end, so links that
# go to nearby places leave from nearby points and cross as little as possible.
left_cursor = {c: left_pos[c][0] for c in left_order}
right_cursor = {c: right_pos[c][0] for c in right_order}
right_rank = {c: i for i, c in enumerate(right_order)}
left_rank = {c: i for i, c in enumerate(left_order)}
for l in sorted(links, key=lambda l: (left_rank[l["origin"]], right_rank[l["destination"]])):
    l["y0"] = left_cursor[l["origin"]]
    left_cursor[l["origin"]] += l["trips"]
for l in sorted(links, key=lambda l: (right_rank[l["destination"]], left_rank[l["origin"]])):
    l["y1"] = right_cursor[l["destination"]]
    right_cursor[l["destination"]] += l["trips"]

layout_height = max(max(b for _, b in left_pos.values()), max(b for _, b in right_pos.values()))

# ---- 5. Write the nodes ------------------------------------------------------------------
name = dict(STATES)
name[NP_CODE] = NP_NAME
nodes = []
for side, pos, size in (("origin", left_pos, left_size), ("destination", right_pos, right_size)):
    for c, (top, bottom) in pos.items():
        nodes.append({
            "side": side, "code": c, "name": name[c],
            "y_top": round(top, 3), "y_bottom": round(bottom, 3),
            "trips": int(round(size[c] * TRIPS_UNIT)),
            # Which origins are np for this destination - for the tooltip. Listed for EVERY
            # destination that has any, including the ACT, whose remainder is not drawn (trap 4).
            "np_origins": ", ".join(np_lists.get(c, [])) if side == "destination" else "",
            "layout_height": round(layout_height, 3),
        })

# ---- 6. Write the links, as curves -------------------------------------------------------
# Each link is a ribbon from (x = 0, its slot on the origin node) to (x = 1, its slot on the
# destination node). Its centre line follows the "smoothstep" curve s(t) = 3t^2 - 2t^3, which
# leaves and arrives HORIZONTALLY - the classic sankey S-shape. Its top and bottom edges follow
# the same curve, so the ribbon keeps the same vertical thickness all the way across.
link_rows = []
dest_total = {d: right_size[d] for d in right_order}
for k, l in enumerate(links):
    link_id = "{0}-{1}".format(l["origin"], l["destination"])
    for i in range(CURVE_SAMPLES + 1):
        t = i / CURVE_SAMPLES
        s = 3 * t ** 2 - 2 * t ** 3
        top = l["y0"] + (l["y1"] - l["y0"]) * s
        link_rows.append({
            "link_id": link_id,
            "origin": l["origin"], "origin_name": name[l["origin"]],
            "destination": l["destination"], "destination_name": name[l["destination"]],
            "t": round(t, 4),
            "y_top": round(top, 3),
            "y_bottom": round(top + l["trips"], 3),
            "trips": int(round(l["trips"] * TRIPS_UNIT)),
            # Share of everything drawn INTO this destination, for the tooltip.
            "share_of_destination": round(l["trips"] / dest_total[l["destination"]], 5),
            "np_origins": l["np_origins"],
        })

os.makedirs("data", exist_ok=True)
for path, records in ((OUT_NODES, nodes), (OUT_LINKS, link_rows)):
    with io.open(path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(records[0].keys()))
        writer.writeheader()
        writer.writerows(records)
    print("wrote {0}  ({1} rows)".format(path, len(records)))

# ---- 7. A report to check the chart against ----------------------------------------------
print("\n{0} links ({1} published, {2} grey np remainders); layout height {3:,.0f} thousand trips".format(
    len(links), sum(l["origin"] != NP_CODE for l in links), sum(l["origin"] == NP_CODE for l in links),
    layout_height))
print("\ninterstate trips drawn INTO each destination (published + np remainder):")
for d in right_order:
    print("  {0:<4} {1:>11,.0f}".format(d, right_size[d] * TRIPS_UNIT))
print("\nnp remainders:")
for d, rem, origins in np_report:
    shown = "{0:>9,.0f} trips".format(rem * TRIPS_UNIT) if rem is not None else "  NOT DRAWN (trap 4)"
    print("  into {0:<4} {1}   np origins: {2}".format(d, shown, ", ".join(origins)))
