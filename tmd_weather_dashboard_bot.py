"""
TMD Phuket Weather -> Raw Excel Log + Grid-Based Atmospheric Dashboard (Single Script).

Excel (Phuket_Weather.xlsx) stores RAW DATA ONLY. The interactive HTML dashboard
(Phuket_Weather_Dashboard.html) features an Apple Weather-inspired grid layout (Date x Time-of-Day),
incorporating dynamic fluid waves, animated wind indicators, atmospheric depth fog, and subtle numeric values.
"""

import argparse
import json
import math
import os
import time
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import requests
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------------

TMD_UID = os.environ.get("TMD_UID", "api")
TMD_UKEY = os.environ.get("TMD_UKEY", "api12345")
TMD_URL = "https://data.tmd.go.th/api/Weather3Hours/V2/"

TARGET_STATION_NAMES = {"PHUKET AIRPORT", "PHUKET"}
TARGET_WMO_NUMBERS = {"48565", "48564"}

VISIBILITY_TAG_CANDIDATES = ["VisibilityLand", "LandVisibility", "Visibility", "Vis"]

OUTPUT_DIR = Path(".")
EXCEL_PATH = OUTPUT_DIR / "Phuket_Weather.xlsx"
HTML_PATH = OUTPUT_DIR / "Phuket_Weather_Dashboard.html"

DATA_SHEET = "Phuket Weather"

LEGACY_SHEET_NAMES_TO_DROP = {"Dashboard", "Lists"}
LEGACY_SHEET_PREFIXES_TO_DROP = ("_pivot_",)

HEADERS = ["Station", "WmoStationNumber", "DateTime", "WindSpeed", "Rainfall_mm", "Rainfall24Hr_mm", "LandVisibility"]

RAINFALL_CAP_MM = 8

STATION_CODES = ["AIRPORT", "PHUKET", "COMBINED"]
STATION_LABELS = {
    "AIRPORT": "PHUKET AIRPORT",
    "PHUKET": "PHUKET",
    "COMBINED": "PHUKET (COMBINED)",
}
STATION_DISPLAY = {
    "AIRPORT": "Phuket Airport",
    "PHUKET": "Phuket",
    "COMBINED": "Phuket (Combined)",
}

METRICS = [
    ("Rain", "Rainfall_mm", "Rainfall (mm)"),
    ("Wind", "WindSpeed", "Wind Speed (km/h)"),
    ("Vis", "LandVisibility", "Land Visibility (km)"),
]

INTERVAL_HOURS = 3

# ----------------------------------------------------------------------------
# Shared Helpers
# ----------------------------------------------------------------------------

def safe_float(value):
    if value in (None, ""):
        return math.nan
    try:
        return float(value)
    except (TypeError, ValueError):
        return math.nan

def none_if_nan(value):
    return None if (value is None or (isinstance(value, float) and math.isnan(value))) else value

def mean_ignore_nan(values):
    present = [v for v in values if not math.isnan(v)]
    if not present:
        return None
    return round(sum(present) / len(present), 2)

def find_text_any(elem, candidates):
    for tag in candidates:
        found = elem.find(tag)
        if found is not None and found.text is not None and found.text.strip() != "":
            return found.text.strip()
    return None

# ----------------------------------------------------------------------------
# Fetch & Parse
# ----------------------------------------------------------------------------

def fetch_raw_records():
    resp = requests.get(TMD_URL, params={"uid": TMD_UID, "ukey": TMD_UKEY}, timeout=30)
    resp.raise_for_status()
    root = ET.fromstring(resp.content)

    records = []
    for station in root.iter("Station"):
        name = (station.findtext("StationNameEnglish") or "").strip()
        wmo = (station.findtext("WmoStationNumber") or "").strip()

        if name.upper() not in TARGET_STATION_NAMES and wmo not in TARGET_WMO_NUMBERS:
            continue

        obs = station.find("Observation")
        if obs is None:
            continue

        dt = (obs.findtext("DateTime") or "").strip()
        wind = safe_float(obs.findtext("WindSpeed"))
        rain_elem = obs.find("Rainfall")
        rain = safe_float(rain_elem.text if rain_elem is not None else None)
        rain24_elem = obs.find("Rainfall24Hr")
        rain24 = safe_float(rain24_elem.text if rain24_elem is not None else None)
        visibility = safe_float(find_text_any(obs, VISIBILITY_TAG_CANDIDATES))

        records.append({
            "station": name.upper(),
            "wmo": wmo,
            "datetime": dt,
            "wind_speed": wind,
            "rainfall_mm": rain,
            "rainfall_24hr_mm": rain24,
            "land_visibility": visibility,
        })
    return records

def build_station_rows(raw_records):
    groups = defaultdict(list)
    for r in raw_records:
        groups[r["datetime"]].append(r)

    rows = []
    for dt in sorted(groups.keys()):
        recs = groups[dt]
        wmo_list = sorted({r["wmo"] for r in recs if r["wmo"]})

        for r in recs:
            rows.append({
                "station": r["station"],
                "wmo": r["wmo"],
                "datetime": dt,
                "wind_speed": none_if_nan(r["wind_speed"]),
                "rainfall_mm": none_if_nan(r["rainfall_mm"]),
                "rainfall_24hr_mm": none_if_nan(r["rainfall_24hr_mm"]),
                "land_visibility": none_if_nan(r["land_visibility"]),
            })

        rows.append({
            "station": STATION_LABELS["COMBINED"],
            "wmo": "+".join(wmo_list),
            "datetime": dt,
            "wind_speed": mean_ignore_nan([r["wind_speed"] for r in recs]),
            "rainfall_mm": mean_ignore_nan([r["rainfall_mm"] for r in recs]),
            "rainfall_24hr_mm": mean_ignore_nan([r["rainfall_24hr_mm"] for r in recs]),
            "land_visibility": mean_ignore_nan([r["land_visibility"] for r in recs]),
        })

    return rows

def fetch_records():
    return build_station_rows(fetch_raw_records())

# ----------------------------------------------------------------------------
# Excel Storage Engine
# ----------------------------------------------------------------------------

def load_or_create_workbook():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if EXCEL_PATH.exists():
        wb = load_workbook(EXCEL_PATH)
        ws = wb[DATA_SHEET] if DATA_SHEET in wb.sheetnames else wb.create_sheet(DATA_SHEET)
        if ws.max_row == 1 and ws.cell(1, 1).value is None:
            ws.append(HEADERS)

        for name in list(wb.sheetnames):
            if name in LEGACY_SHEET_NAMES_TO_DROP or name.startswith(LEGACY_SHEET_PREFIXES_TO_DROP):
                del wb[name]
        wb.active = wb.sheetnames.index(DATA_SHEET)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = DATA_SHEET
        ws.append(HEADERS)

    for col_idx in range(1, len(HEADERS) + 1):
        ws.cell(1, col_idx).font = Font(bold=True)

    return wb, ws

def update_excel(records):
    wb, ws = load_or_create_workbook()

    existing = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        existing.add((row[0], row[2]))

    added = 0
    for r in records:
        key = (r["station"], r["datetime"])
        if key in existing:
            continue
        ws.append([r["station"], r["wmo"], r["datetime"], r["wind_speed"], r["rainfall_mm"], r["rainfall_24hr_mm"], r["land_visibility"]])
        existing.add(key)
        added += 1

    for col_idx in range(1, len(HEADERS) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 20

    wb.save(EXCEL_PATH)
    return added, ws.max_row - 1

# ----------------------------------------------------------------------------
# HTML Dashboard Generator
# ----------------------------------------------------------------------------

def generate_html_dashboard(html_path=HTML_PATH):
    config_json = json.dumps({
        "data_sheet_name": DATA_SHEET,
        "xlsx_filename": EXCEL_PATH.name,
        "metrics": [{"code": c, "column": col, "label": label} for c, col, label in METRICS],
        "stations": [
            {"code": c, "label_raw": STATION_LABELS[c], "label_display": STATION_DISPLAY[c]}
            for c in STATION_CODES
        ],
        "rainfall_cap": RAINFALL_CAP_MM,
    })
    html = _HTML_TEMPLATE.replace("__CONFIG_JSON__", config_json)
    html_path.parent.mkdir(parents=True, exist_ok=True)
    html_path.write_text(html, encoding="utf-8")


_HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Meteorological Intelligence | Phuket</title>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
  :root {
    --bg-color: #F8FAFC;
    --card-bg: rgba(255, 255, 255, 0.85);
    --card-border: rgba(226, 232, 240, 0.8);
    --glass-shadow: 0 10px 30px -10px rgba(0, 0, 0, 0.05);
    --primary-blue: #3A6FF7;
    --text-primary: #0F172A;
    --text-muted: #64748B;
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "SF Pro Text", "Helvetica Neue", sans-serif;
    background-color: var(--bg-color);
    color: var(--text-primary);
    padding: 24px;
    max-width: 1300px;
    margin: 0 auto;
    -webkit-font-smoothing: antialiased;
  }

  header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 24px;
    flex-wrap: wrap;
    gap: 16px;
  }

  h1 {
    font-size: 26px;
    font-weight: 600;
    letter-spacing: -0.5px;
  }

  .subtitle {
    font-size: 13px;
    color: var(--text-muted);
    margin-top: 2px;
  }

  .controls {
    display: flex;
    gap: 12px;
    align-items: center;
  }

  select {
    appearance: none;
    background: var(--card-bg);
    border: 1px solid var(--card-border);
    padding: 8px 16px;
    border-radius: 12px;
    font-size: 13px;
    font-weight: 500;
    color: var(--text-primary);
    outline: none;
    cursor: pointer;
    box-shadow: var(--glass-shadow);
  }

  .grid-container {
    background: var(--card-bg);
    border: 1px solid var(--card-border);
    border-radius: 20px;
    backdrop-filter: blur(20px);
    box-shadow: var(--glass-shadow);
    overflow: auto;
    max-height: 75vh;
  }

  table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
    text-align: center;
  }

  th, td {
    border-bottom: 1px solid var(--card-border);
    border-right: 1px solid var(--card-border);
    padding: 0;
    min-width: 72px;
    height: 60px;
  }

  th {
    background: #F1F5F9;
    color: var(--text-muted);
    font-weight: 600;
    font-size: 12px;
    position: sticky;
    top: 0;
    z-index: 10;
    height: 40px;
  }

  td.datecol, th.datecol {
    position: sticky;
    left: 0;
    background: #F1F5F9;
    z-index: 11;
    font-weight: 600;
    text-align: left;
    padding: 0 16px;
    min-width: 110px;
    color: var(--text-primary);
  }

  td.summary {
    background: rgba(58, 111, 247, 0.08);
    font-weight: 700;
    color: var(--primary-blue);
    vertical-align: middle;
  }

  .cell-wrapper {
    position: relative;
    width: 100%;
    height: 100%;
    display: flex;
    align-items: center;
    justify-content: center;
    overflow: hidden;
  }

  .wave-bg {
    position: absolute;
    left: 0; right: 0; bottom: 0;
    background: linear-gradient(180deg, rgba(58, 111, 247, 0.4) 0%, rgba(58, 111, 247, 0.7) 100%);
    transition: height 0.3s ease;
  }

  .wind-icon {
    width: 16px;
    height: 16px;
    fill: #3A6FF7;
    margin-right: 4px;
    display: inline-block;
    animation: spin linear infinite;
  }
  @keyframes spin { 100% { transform: rotate(360deg); } }

  .fog-layer {
    position: absolute;
    inset: 0;
    background: rgba(255, 255, 255, 0.6);
    backdrop-filter: blur(var(--fog-blur));
    pointer-events: none;
  }

  .cell-val {
    position: relative;
    z-index: 2;
    font-size: 12px;
    font-weight: 600;
    font-variant-numeric: tabular-nums;
    color: var(--text-primary);
  }
  .empty-val { color: #CBD5E1; }
</style>
</head>
<body>

  <header>
    <div>
      <h1>Phuket Weather Intelligence</h1>
      <div class="subtitle" id="timestamp">Syncing live data...</div>
    </div>
    <div class="controls">
      <select id="metricSel"></select>
      <select id="stationSel"></select>
    </div>
  </header>

  <div class="grid-container">
    <table id="gridTable"></table>
  </div>

<script>
const CONFIG = __CONFIG_JSON__;
const metricSel = document.getElementById('metricSel');
const stationSel = document.getElementById('stationSel');
const table = document.getElementById('gridTable');
let DATA = null;

CONFIG.metrics.forEach(m => {
  const opt = document.createElement('option');
  opt.value = m.code; opt.textContent = m.label;
  metricSel.appendChild(opt);
});

CONFIG.stations.forEach(s => {
  const opt = document.createElement('option');
  opt.value = s.code; opt.textContent = s.label_display;
  if (s.code === 'COMBINED') opt.selected = true;
  stationSel.appendChild(opt);
});

async function autoFetchServerData() {
  try {
    const res = await fetch(CONFIG.xlsx_filename, { cache: 'no-cache' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const buffer = await res.arrayBuffer();
    
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
    const sheet = wb.Sheets[CONFIG.data_sheet_name];
    if (!sheet) throw new Error("Data sheet missing");

    const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });
    DATA = buildDataFromRows(rows);
    document.getElementById('timestamp').textContent = `Updated • ${new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}`;
    renderGrid();
  } catch (err) {
    document.getElementById('timestamp').textContent = 'Displaying static cache';
  }
}

function pad(n) { return n.toString().padStart(2, '0'); }
function parseDateTime(v) {
  if (!v) return null;
  const d = (v instanceof Date) ? v : new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function buildDataFromRows(rows) {
  const buckets = {};
  const dateSet = new Set();
  const timeSet = new Set();
  const columns = CONFIG.metrics.map(m => m.column);

  rows.forEach(row => {
    const dt = parseDateTime(row.DateTime);
    const station = row.Station;
    if (!dt || !station) return;

    const dateStr = `${dt.getFullYear()}-${pad(dt.getMonth() + 1)}-${pad(dt.getDate())}`;
    const timeStr = `${pad(dt.getHours())}:${pad(dt.getMinutes())}`;
    dateSet.add(dateStr); timeSet.add(timeStr);

    buckets[station] = buckets[station] || {};
    buckets[station][dateStr] = buckets[station][dateStr] || {};
    const slot = buckets[station][dateStr][timeStr] =
      buckets[station][dateStr][timeStr] || Object.fromEntries(columns.map(c => [c, []]));

    columns.forEach(col => {
      const v = row[col];
      if (v !== null && v !== undefined && v !== '' && !isNaN(Number(v))) slot[col].push(Number(v));
    });
  });

  const dates = Array.from(dateSet).sort();
  const times = Array.from(timeSet).sort((a, b) => a.localeCompare(b));

  const round2 = x => Math.round(x * 100) / 100;
  const mean = arr => arr.length ? round2(arr.reduce((a, b) => a + b, 0) / arr.length) : null;

  const payload = {};
  CONFIG.metrics.forEach(m => {
    payload[m.code] = {};
    CONFIG.stations.forEach(s => {
      const stationLabel = s.label_raw;
      const grid = dates.map(d => times.map(t => {
        const slot = buckets[stationLabel] && buckets[stationLabel][d] && buckets[stationLabel][d][t];
        return slot ? mean(slot[m.column]) : null;
      }));
      const summary = grid.map(rowVals => {
        const present = rowVals.filter(v => v !== null);
        if (!present.length) return null;
        return m.code === 'Rain' ? round2(present.reduce((a, b) => a + b, 0)) : mean(present);
      });
      payload[m.code][s.code] = { grid, summary };
    });
  });

  return { dates, times, payload };
}

function renderGrid() {
  if (!DATA) return;

  const metric = metricSel.value;
  const station = stationSel.value;
  const combo = DATA.payload[metric][station];
  const grid = combo.grid, summary = combo.summary;

  table.innerHTML = '';

  // Header Row
  const thead = document.createElement('tr');
  thead.appendChild(Object.assign(document.createElement('th'), {textContent: 'Date', className: 'datecol'}));
  DATA.times.forEach(t => thead.appendChild(Object.assign(document.createElement('th'), {textContent: t})));
  thead.appendChild(Object.assign(document.createElement('th'), {textContent: 'Summary'}));
  table.appendChild(thead);

  // Data Rows
  DATA.dates.forEach((d, i) => {
    const tr = document.createElement('tr');

    const dateTd = document.createElement('td');
    dateTd.textContent = d; dateTd.className = 'datecol';
    tr.appendChild(dateTd);

    grid[i].forEach(v => {
      const td = document.createElement('td');
      const wrapper = document.createElement('div');
      wrapper.className = 'cell-wrapper';

      if (metric === 'Rain') {
        if (v !== null && v > 0) {
          const fillRatio = Math.min(v, CONFIG.rainfall_cap) / CONFIG.rainfall_cap;
          const wave = document.createElement('div');
          wave.className = 'wave-bg';
          wave.style.height = (fillRatio * 100) + '%';
          wrapper.appendChild(wave);
        }
        const valSpan = document.createElement('span');
        valSpan.className = 'cell-val' + (v === null ? ' empty-val' : '');
        valSpan.textContent = v === null ? '-' : v;
        wrapper.appendChild(valSpan);

      } else if (metric === 'Wind') {
        if (v !== null && v > 0) {
          const spinSpeed = Math.max(0.2, 3 - (v / 10));
          wrapper.innerHTML = `
            <svg class="wind-icon" style="animation-duration: ${spinSpeed}s" viewBox="0 0 24 24">
              <path d="M12,11A1,1 0 0,0 11,12A1,1 0 0,0 12,13A1,1 0 0,0 13,12A1,1 0 0,0 12,11M12,2A10,10 0 0,0 2,12A10,10 0 0,0 12,22A10,10 0 0,0 22,12A10,10 0 0,0 12,2M12,4C13.2,4 14.18,4.86 14.36,6C13.5,6 12.7,6.36 12.12,6.94C11.54,7.5 11.18,8.3 11.18,9.18C10.3,9.18 9.5,9.54 8.92,10.12C8.34,10.7 8,11.5 8,12.38C8,13.56 8.86,14.54 10,14.72C10,15.58 10.36,16.38 10.94,16.96C11.5,17.54 12.3,17.9 13.18,17.9C14.36,17.9 15.34,17.04 15.52,15.88C16.4,15.88 17.2,15.5 17.78,14.94C18.36,14.36 18.72,13.56 18.72,12.68C18.72,11.5 17.86,10.5 16.7,10.34C16.7,9.46 16.34,8.66 15.76,8.08C15.18,7.5 14.38,7.14 13.5,7.14C13.5,5.96 12.64,5 11.5,5.18C11.66,4.5 12,4 12,4Z"/>
            </svg>
            <span class="cell-val">${v}</span>`;
        } else {
          wrapper.innerHTML = `<span class="cell-val ${v === null ? 'empty-val' : ''}">${v === null ? '-' : v}</span>`;
        }

      } else if (metric === 'Vis') {
        if (v !== null) {
          if (v < 10) {
            const fog = document.createElement('div');
            fog.className = 'fog-layer';
            fog.style.setProperty('--fog-blur', `${Math.max(1, (10 - v) * 0.8)}px`);
            wrapper.appendChild(fog);
          }
          const valSpan = document.createElement('span');
          valSpan.className = 'cell-val';
          valSpan.textContent = v;
          wrapper.appendChild(valSpan);
        } else {
          wrapper.innerHTML = `<span class="cell-val empty-val">-</span>`;
        }
      }

      td.appendChild(wrapper);
      tr.appendChild(td);
    });

    const sumTd = document.createElement('td');
    sumTd.className = 'summary';
    sumTd.textContent = summary[i] === null ? '-' : summary[i];
    tr.appendChild(sumTd);

    table.appendChild(tr);
  });
}

metricSel.addEventListener('change', renderGrid);
stationSel.addEventListener('change', renderGrid);
window.addEventListener('DOMContentLoaded', autoFetchServerData);
</script>
</body>
</html>
"""

# ----------------------------------------------------------------------------
# Orchestration
# ----------------------------------------------------------------------------

def run_cycle():
    print(f"[{datetime.now().isoformat()}] Fetching TMD feed...")
    records = fetch_records()
    print(f"Built {len(records)} row(s) (raw + combined) this cycle.")

    added, total = update_excel(records)
    print(f"Excel Data sheet updated: +{added} new row(s), {total} total.")

    generate_html_dashboard(HTML_PATH)
    print(f"Atmospheric HTML dashboard generated at {HTML_PATH}")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--once", action="store_true", help="Run a single cycle and exit")
    args = parser.parse_args()

    if args.once:
        run_cycle()
        return

    while True:
        try:
            run_cycle()
        except Exception as e:
            print(f"Error this cycle: {e}")
        print(f"Sleeping {INTERVAL_HOURS} hour(s)...")
        time.sleep(INTERVAL_HOURS * 3600)

if __name__ == "__main__":
    main()
