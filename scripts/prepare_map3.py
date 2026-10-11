"""Map 3 preprocessing: pip install shapely; python prepare_map3.py.

All terrestrial governance types; mainland/Tasmania and nearby islands only.
Count distinct PA_IDs from record centroids per 200 km spherical Albers cell.
An ID may appear in multiple cells. Grey means no recorded centroid, not
unprotected land. Counts use full cells; coast clipping affects display only.
Underlying cells are equal-area; visible coastal fragments are not.
The authalic sphere matches Vega conicEqualArea, not ellipsoidal EPSG:3577.
"""
import json
import math
from collections import defaultdict
from pathlib import Path
from shapely import make_valid, segmentize
from shapely.geometry import Polygon, MultiPolygon, box, mapping
from shapely.geometry.polygon import orient
from shapely.ops import transform, unary_union

ROOT = Path(__file__).resolve().parent.parent   # scripts/ sits one level inside the project
BIN_KM = 200
RADIUS_KM = 6371.0071809
SOURCE = ROOT / "raw-data" / "Collaborative_Australian_Protected_Areas_Database_(CAPAD)_2024_-_Terrestrial__.json"
PHI1, PHI2 = map(math.radians, (-18, -36))
N = (math.sin(PHI1) + math.sin(PHI2)) / 2
C = math.cos(PHI1) ** 2 + 2 * N * math.sin(PHI1)
RHO0 = RADIUS_KM * math.sqrt(C) / N


def project(lon, lat):
    theta = N * math.radians(lon - 132)
    rho = RADIUS_KM * math.sqrt(C - 2 * N * math.sin(math.radians(lat))) / N
    return rho * math.sin(theta), RHO0 - rho * math.cos(theta)


def unproject(x, y):
    rho = -math.hypot(x, RHO0 - y)
    theta = math.atan2(-x, -(RHO0 - y))
    lat = math.asin((C - (rho * N / RADIUS_KM) ** 2) / (2 * N))
    return round(132 + math.degrees(theta / N), 6), round(math.degrees(lat), 6)


def read_land():
    """Decode the existing TopoJSON coastline and union state polygons."""
    topo = json.loads((ROOT / "docs" / "data" / "Australia_topo.json").read_text())
    arcs = []
    for arc in topo["arcs"]:
        points, x, y = [], 0, 0
        for a, b in arc:
            if "transform" in topo:
                x, y = x + a, y + b
                lon = x * topo["transform"]["scale"][0] + topo["transform"]["translate"][0]
                lat = y * topo["transform"]["scale"][1] + topo["transform"]["translate"][1]
            else:
                lon, lat = a, b
            points.append(project(lon, lat))
        arcs.append(points)

    def ring(indices):
        result = []
        for index in indices:
            points = arcs[index] if index >= 0 else arcs[~index][::-1]
            result.extend(points if not result else points[1:])
        return result

    polygons = []
    for g in topo["objects"]["map"]["geometries"]:
        if g["properties"]["name"] in ("Macquarie Island", "Norfolk Island"):
            continue
        parts = [g["arcs"]] if g["type"] == "Polygon" else g["arcs"]
        for rings in parts:
            polygons.append(make_valid(Polygon(ring(rings[0]), [ring(r) for r in rings[1:]])))
    return unary_union(polygons)


def polygons_only(geometry):
    if geometry.is_empty:
        return []
    if geometry.geom_type == "Polygon":
        return [geometry]
    return [p for g in getattr(geometry, "geoms", []) for p in polygons_only(g)]


def geographic(geometry):
    # Densify edges before inverse projection; D3 expects clockwise exteriors.
    geometry = segmentize(geometry, max_segment_length=10)
    if geometry.geom_type in ("Polygon", "MultiPolygon"):
        parts = [orient(p, sign=-1) for p in polygons_only(geometry)]
        geometry = parts[0] if len(parts) == 1 else MultiPolygon(parts)
    return mapping(transform(unproject, geometry))


def main():
    land = read_land()
    bins = defaultdict(lambda: defaultdict(set))
    topo = json.loads(SOURCE.read_text(encoding="utf-8"))
    excluded = 0
    governance_codes = set()
    for geometry in next(iter(topo["objects"].values()))["geometries"]:
        p = geometry.get("properties") or {}
        governance = p.get("GOVERNANCE") or "Unknown"
        governance_codes.add(governance)
        lon, lat = float(p["LONGITUDE"]), float(p["LATITUDE"])
        if not (110 <= lon <= 155 and -45 <= lat <= -9):
            excluded += 1
            continue
        pa_id = str(p.get("PA_ID") or "").strip()
        if not pa_id:
            raise ValueError(f"Missing PA_ID: {p.get('NAME')}")
        x, y = project(lon, lat)
        bins[math.floor(x / BIN_KM), math.floor(y / BIN_KM)][governance].add(pa_id)

    features, displayed = [], set()
    minx, miny, maxx, maxy = land.bounds
    for col in range(math.floor(minx / BIN_KM), math.floor(maxx / BIN_KM) + 1):
        for row in range(math.floor(miny / BIN_KM), math.floor(maxy / BIN_KM) + 1):
            x, y = col * BIN_KM, row * BIN_KM
            parts = polygons_only(box(x, y, x + BIN_KM, y + BIN_KM).intersection(land))
            if not parts:
                continue
            clipped = unary_union(parts)
            groups = bins[col, row]
            all_ids = set().union(*groups.values())
            props = {
                "bin_id": f"{col}:{row}", "bin_size": f"{BIN_KM} x {BIN_KM} km",
                "distinct_areas": len(all_ids),
                # Keep every observed code separate, including unknown/new codes.
                "governance_counts": {code: len(groups[code]) for code in sorted(governance_codes)},
            }
            features.append({"type": "Feature", "properties": props, "geometry": geographic(clipped)})
            displayed.add((col, row))

    missing = [key for key, groups in bins.items() if any(groups.values()) and key not in displayed]
    # The simplified basemap omits some small offshore islands. Do not move
    # their centroids onto mainland cells or invent a coastline for them.
    omitted_ids = sorted(set().union(*(set().union(*bins[key].values()) for key in missing)))
    metadata = {
        "scope": "Terrestrial; mainland Australia, Tasmania and islands in the supplied basemap",
        "bin_km": BIN_KM, "outside_view_records": excluded,
        "governance_codes": sorted(governance_codes),
        "ids_without_basemap_land": omitted_ids,
        "count_method": "Distinct PA_ID per full cell; coastal clipping is display-only",
    }
    output = ROOT / "docs" / "data" / "capad2024_governance_bins.geojson"
    output.write_text(json.dumps({"type": "FeatureCollection", "metadata": metadata, "features": features}, separators=(",", ":")) + "\n")
    print(f"{len(features)} land cells; {sum(f['properties']['distinct_areas'] == 0 for f in features)} empty; "
          f"governance categories: {', '.join(sorted(governance_codes))}; "
          f"{excluded} records outside view; {len(omitted_ids)} island IDs absent from basemap; "
          f"{output.stat().st_size:,} bytes")


if __name__ == "__main__":
    main()
