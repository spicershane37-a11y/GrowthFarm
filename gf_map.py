# gf_map.py
# Build & open a simple Leaflet map from customers.csv (and customers_geo.csv sidecar).

import csv
from pathlib import Path

from gf_store import APP_DIR, CUSTOMERS_PATH

def _read_geo_sidecar():
    """Return (by_company, by_addrkey) from customers_geo.csv if present."""
    path = APP_DIR / "customers_geo.csv"
    by_company, by_addr = {}, {}
    if not path.exists():
        return by_company, by_addr
    try:
        with path.open("r", encoding="utf-8", newline="") as f:
            for r in csv.DictReader(f):
                try:
                    la = float((r.get("Lat") or "").strip())
                    lo = float((r.get("Lon") or "").strip())
                except Exception:
                    continue
                comp  = (r.get("Company") or "").strip().lower()
                addrk = (r.get("AddressKey") or "").strip().lower()
                if comp:  by_company[comp] = (la, lo)
                if addrk: by_addr[addrk]   = (la, lo)
    except Exception:
        pass
    return by_company, by_addr

def _addr_key_from_row(r):
    addr = (r.get("Address") or "").strip()
    city = (r.get("City") or "").strip()
    state = (r.get("State") or "").strip()
    zipc = (r.get("ZIP") or "").strip()
    if any([addr, city, state, zipc]):
        parts = [p for p in (addr, city, state, zipc) if (p or "").strip()]
        return ", ".join(parts).lower().strip()
    return (r.get("Location") or "").strip().lower()

def _money_fmt(s):
    s = (str(s or "").replace("$","").replace(",","").strip())
    if not s: return ""
    try:
        return f"${float(s):,.2f}"
    except Exception:
        return s

def _load_customers_for_map():
    """Return (records, skipped_without_coords)."""
    recs, skipped = [], 0
    by_company, by_addr = _read_geo_sidecar()

    if not CUSTOMERS_PATH.exists():
        return recs, skipped

    with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
        for r in csv.DictReader(f):
            # prefer explicit Lat/Lon in CSV
            lat_s = (r.get("Lat") or r.get("Latitude") or "").strip()
            lon_s = (r.get("Lon") or r.get("Lng") or r.get("Longitude") or "").strip()
            lat = lon = None
            if lat_s and lon_s:
                try:
                    lat, lon = float(lat_s), float(lon_s)
                except Exception:
                    lat = lon = None

            # else try sidecar by Company, then by address key
            if lat is None or lon is None:
                comp_key = (r.get("Company") or "").strip().lower()
                addr_key = _addr_key_from_row(r)
                hit = by_company.get(comp_key) or by_addr.get(addr_key)
                if hit:
                    lat, lon = hit

            if lat is None or lon is None:
                skipped += 1
                continue

            company = (r.get("Company") or "(Unnamed)").strip()
            cltv    = _money_fmt(r.get("CLTV"))
            spd_raw = (r.get("Sales/Day") or r.get("Sales per Day") or "").strip()
            spd     = _money_fmt(spd_raw) if spd_raw else "—"

            popup = (
                f"<b>{company}</b><br/>"
                f"CLTV: {cltv or '$0.00'}<br/>"
                f"Sales/Day: {spd}"
            )
            recs.append({"lat": lat, "lon": lon, "popup": popup})
    return recs, skipped

def _write_leaflet_html(recs, out_path: Path):
    if recs:
        try:
            avg_lat = sum(r["lat"] for r in recs) / len(recs)
            avg_lon = sum(r["lon"] for r in recs) / len(recs)
        except Exception:
            avg_lat, avg_lon = 39.5, -98.35
    else:
        avg_lat, avg_lon = 39.5, -98.35

    markers = []
    for r in recs:
        p = r["popup"].replace("\\", "\\\\").replace("`", "\\`")
        markers.append(
            f"L.marker([{r['lat']:.6f}, {r['lon']:.6f}]).addTo(map).bindPopup(`{p}`);"
        )
    markers_js = "\n    ".join(markers)

    html = f"""<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Customer Map</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<style>
  html, body, #map {{ height: 100%; margin: 0; background:#111; }}
  .leaflet-popup-content-wrapper, .leaflet-popup-tip {{ background:#222; color:#eee; }}
</style>
</head>
<body>
<div id="map"></div>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script>
  var map = L.map('map').setView([{avg_lat:.6f}, {avg_lon:.6f}], {12 if len(recs)==1 else 5});
  L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
    maxZoom: 19,
    attribution: '&copy; OpenStreetMap'
  }}).addTo(map);
  {markers_js}
</script>
</body>
</html>"""
    out_path.write_text(html, encoding="utf-8")

def open_customer_map(window=None):
    """Build map HTML and open it. Updates -MAP_STATUS- label if provided."""
    out_path = APP_DIR / "customer_map.html"
    recs, skipped = _load_customers_for_map()
    if not recs and skipped == 0:
        if window and "-MAP_STATUS-" in getattr(window, "AllKeysDict", {}):
            window["-MAP_STATUS-"].update("No customers yet.")
        return

    try:
        _write_leaflet_html(recs, out_path)
        import webbrowser
        webbrowser.open(str(out_path))
        msg = f"Opened map ({len(recs)} pin(s){', skipped ' + str(skipped) + ' without coords' if skipped else ''})."
    except Exception as e:
        msg = f"Map error: {e}"

    if window and "-MAP_STATUS-" in getattr(window, "AllKeysDict", {}):
        window["-MAP_STATUS-"].update(msg)
