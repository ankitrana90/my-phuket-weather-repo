"""
TMD Phuket Weather -> raw Excel log + dynamic HTML dashboard (single script).

Excel (Phuket_Weather.xlsx) stores RAW DATA ONLY. All dynamic/interactive views live in
Phuket_Weather_Dashboard.html, which automatically fetches the xlsx file directly over HTTP(S) when opened.
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
# Shared helpers
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
# Excel Logic
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
<title>Phuket Weather Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
  :root {
    --bg: #0b0f19;
    --card-bg: rgba(22, 31, 49, 0.75);
    --border: rgba(255, 255, 255, 0.08);
    --text: #f3f4f6;
    --text-muted: #9ca3af;
    --accent: #3b82f6;
    --summary-bg: rgba(59, 130, 246, 0.12);
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'Plus Jakarta Sans', sans-serif;
    background: var(--bg);
    color: var(--text);
    padding: 28px;
    min-height: 100vh;
    background-image: 
      radial-gradient(at 0% 0%, rgba(59, 130, 246, 0.12) 0px, transparent 50%),
      radial-gradient(at 100% 100%, rgba(14, 165, 233, 0.1) 0px, transparent 50%);
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
    font-size: 24px;
    font-weight: 700;
    letter-spacing: -0.5px;
    background: linear-gradient(135deg, #fff 0%, #94a3b8 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
  }

  #fileStatus {
    font-size: 13px;
    color: var(--text-muted);
    background: rgba(255, 255, 255, 0.04);
    padding: 6px 14px;
    border-radius: 20px;
    border: 1px solid var(--border);
    display: inline-flex;
    align-items: center;
    gap: 8px;
  }
  #fileStatus.ok { color: #10b981; border-color: rgba(16, 185, 129, 0.2); }
  #fileStatus.error { color: #ef4444; border-color: rgba(239, 68, 68, 0.2); }

  .controls-bar {
    display: flex;
    gap: 20px;
    background: var(--card-bg);
    padding: 16px 20px;
    border-radius: 14px;
    border: 1px solid var(--border);
    backdrop-filter: blur(12px);
    margin-bottom: 20px;
    align-items: center;
    flex-wrap: wrap;
  }

  .control-group {
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .control-group label {
    font-size: 13px;
    font-weight: 600;
    color: var(--text-muted);
    text-transform: uppercase;
    letter-spacing: 0.5px;
  }

  select {
    background: #111827;
    color: var(--text);
    border: 1px solid var(--border);
    padding: 8px 14px;
    border-radius: 8px;
    font-size: 14px;
    font-family: inherit;
    outline: none;
    cursor: pointer;
    transition: all 0.2s;
  }
  select:hover { border-color: var(--accent); }

  .wrap {
    background: var(--card-bg);
    border-radius: 16px;
    border: 1px solid var(--border);
    backdrop-filter: blur(12px);
    overflow: auto;
    max-height: 72vh;
    box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.3);
  }

  table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
    text-align: center;
  }

  th, td {
    padding: 10px 12px;
    border-bottom: 1px solid var(--border);
    border-right: 1px solid var(--border);
    min-width: 68px;
  }

  th {
    background: #111827;
    color: var(--text-muted);
    font-weight: 600;
    position: sticky;
    top: 0;
    z-index: 10;
  }

  td.datecol, th.datecol {
    position: sticky;
    left: 0;
    background: #111827;
    z-index: 11;
    font-weight: 600;
    text-align: left;
    min-width: 110px;
    color: var(--text);
  }

  td.summary {
    background: var(--summary-bg);
    font-weight: 700;
    color: #60a5fa;
  }

  /* ---------------------------------------------------
     METRIC VISUALIZATIONS
  --------------------------------------------------- */

  /* 1. Rainfall Liquid Wave Container */
  td.rain-cell {
    position: relative;
    height: 64px;
    padding: 0;
    vertical-align: bottom;
    overflow: hidden;
    background: rgba(15, 23, 42, 0.6);
  }

  .wave-bucket {
    position: absolute;
    left: 0; right: 0; bottom: 0;
    background: linear-gradient(180deg, #3b82f6 0%, #1d4ed8 100%);
    transition: height 0.4s ease-out;
  }

  /* Wave Animation effect */
  .wave-bucket::before {
    content: "";
    position: absolute;
    top: -8px; left: 0; width: 200%; height: 12px;
    background: repeating-linear-gradient(90deg, rgba(255,255,255,0.4) 0px, transparent 12px, rgba(255,255,255,0.4) 24px);
    animation: waveMove 2s infinite linear;
  }

  @keyframes waveMove {
    0% { transform: translateX(0); }
    100% { transform: translateX(-50%); }
  }

  td.rain-cell .val {
    position: relative;
    z-index: 2;
    font-weight: 700;
    display: block;
    line-height: 64px;
    text-shadow: 0 1px 3px rgba(0,0,0,0.8);
  }

  /* 2. Wind Speed Fan Animation */
  td.wind-cell {
    height: 64px;
    vertical-align: middle;
  }
  .wind-container {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
  }
  .fan-icon {
    width: 18px;
    height: 18px;
    fill: #38bdf8;
    display: inline-block;
    animation: spin linear infinite;
  }
  @keyframes spin {
    100% { transform: rotate(360deg); }
  }

  /* 3. Frost Visibility Effect */
  td.vis-cell {
    height: 64px;
    vertical-align: middle;
    position: relative;
    overflow: hidden;
  }
  .frost-overlay {
    position: absolute;
    inset: 0;
    background: rgba(255, 255, 255, 0.25);
    backdrop-filter: blur(var(--blur-amount));
    pointer-events: none;
  }

  .empty { color: rgba(255,255,255,0.2); }
</style>
</head>
<body>

  <header>
    <div>
      <h1>Phuket Weather Dashboard</h1>
    </div>
    <div id="fileStatus">Initializing live data connection...</div>
  </header>

  <div class="controls-bar">
    <div class="control-group">
      <label for="metric">Metric</label>
      <select id="metric"></select>
    </div>
    <div class="control-group">
      <label for="station">Station</label>
      <select id="station"></select>
    </div>
  </div>

  <div class="wrap"><table id="grid"></table></div>

<script>
const CONFIG = __CONFIG_JSON__;
const RAIN_CAP = CONFIG.rainfall_cap;

const metricSel = document.getElementById('metric');
const stationSel = document.getElementById('station');
const statusEl = document.getElementById('fileStatus');
const table = document.getElementById('grid');

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

function setStatus(msg, cls) {
  statusEl.textContent = msg;
  statusEl.className = cls || '';
}

async function autoFetchServerData() {
  setStatus('Syncing live data...', '');
  try {
    const res = await fetch(CONFIG.xlsx_filename, { cache: 'no-cache' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const buffer = await res.arrayBuffer();
    
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
    const sheet = wb.Sheets[CONFIG.data_sheet_name];
    if (!sheet) throw new Error(`Sheet "${CONFIG.data_sheet_name}" missing`);

    const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });
    DATA = buildDataFromRows(rows);
    setStatus(`Updated live: ${new Date().toLocaleTimeString()}`, 'ok');
    render();
  } catch (err) {
    setStatus('Unable to load server dataset', 'error');
  }
}

function pad(n) { return n.toString().padStart(2, '0'); }

function parseDateTime(v) {
  if (v === null || v === undefined || v === '') return null;
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
  const times = Array.from(timeSet).sort((a, b) => {
    const [ah, am] = a.split(':').map(Number), [bh, bm] = b.split(':').map(Number);
    return (ah * 60 + am) - (bh * 60 + bm);
  });

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

function render() {
  if (!DATA) return;

  const metric = metricSel.value;
  const station = stationSel.value;
  const combo = DATA.payload[metric][station];
  const grid = combo.grid, summary = combo.summary;

  table.innerHTML = '';

  const thead = document.createElement('tr');
  thead.appendChild(Object.assign(document.createElement('th'), {textContent: 'Date', className: 'datecol'}));
  DATA.times.forEach(t => thead.appendChild(Object.assign(document.createElement('th'), {textContent: t})));
  thead.appendChild(Object.assign(document.createElement('th'), {textContent: 'Daily Summary'}));
  table.appendChild(thead);

  DATA.dates.forEach((d, i) => {
    const tr = document.createElement('tr');
    const dateTd = document.createElement('td');
    dateTd.textContent = d; dateTd.className = 'datecol';
    tr.appendChild(dateTd);

    grid[i].forEach(v => {
      const td = document.createElement('td');

      if (metric === 'Rain') {
        td.className = 'rain-cell';
        if (v !== null && v > 0) {
          const fillRatio = Math.min(v, RAIN_CAP) / RAIN_CAP;
          const bucket = document.createElement('div');
          bucket.className = 'wave-bucket';
          bucket.style.height = (fillRatio * 100) + '%';
          td.appendChild(bucket);
        }
        const span = document.createElement('span');
        span.className = 'val';
        span.textContent = v === null ? '' : v;
        if (v === null) span.className += ' empty';
        td.appendChild(span);

      } else if (metric === 'Wind') {
        td.className = 'wind-cell';
        if (v !== null && v > 0) {
          const spinDuration = Math.max(0.2, 3 - (v / 10)); // Speed scales fan animation
          td.innerHTML = `
            <div class="wind-container">
              <svg class="fan-icon" style="animation-duration: ${spinDuration}s" viewBox="0 0 24 24">
                <path d="M12,11A1,1 0 0,0 11,12A1,1 0 0,0 12,13A1,1 0 0,0 13,12A1,1 0 0,0 12,11M12,2A10,10 0 0,0 2,12A10,10 0 0,0 12,22A10,10 0 0,0 22,12A10,10 0 0,0 12,2M12,4C13.2,4 14.18,4.86 14.36,6C13.5,6 12.7,6.36 12.12,6.94C11.54,7.5 11.18,8.3 11.18,9.18C10.3,9.18 9.5,9.54 8.92,10.12C8.34,10.7 8,11.5 8,12.38C8,13.56 8.86,14.54 10,14.72C10,15.58 10.36,16.38 10.94,16.96C11.5,17.54 12.3,17.9 13.18,17.9C14.36,17.9 15.34,17.04 15.52,15.88C16.4,15.88 17.2,15.5 17.78,14.94C18.36,14.36 18.72,13.56 18.72,12.68C18.72,11.5 17.86,10.5 16.7,10.34C16.7,9.46 16.34,8.66 15.76,8.08C15.18,7.5 14.38,7.14 13.5,7.14C13.5,5.96 12.64,5 11.5,5.18C11.66,4.5 12,4 12,4Z"/>
              </svg>
              <span>${v}</span>
            </div>`;
        } else {
          td.textContent = v === null ? '' : v;
          if (v === null) td.className += ' empty';
        }

      } else if (metric === 'Vis') {
        td.className = 'vis-cell';
        if (v !== null) {
          // Low visibility creates heavier frost
          if (v < 10) {
            const blurAmount = Math.max(1, (10 - v) * 0.8);
            const overlay = document.createElement('div');
            overlay.className = 'frost-overlay';
            overlay.style.setProperty('--blur-amount', `${blurAmount}px`);
            td.appendChild(overlay);
          }
          const span = document.createElement('span');
          span.style.position = 'relative';
          span.style.zIndex = '2';
          span.textContent = v;
          td.appendChild(span);
        } else {
          td.className += ' empty';
        }
      }

      tr.appendChild(td);
    });

    const sumTd = document.createElement('td');
    sumTd.className = 'summary';
    sumTd.textContent = summary[i] === null ? '' : summary[i];
    tr.appendChild(sumTd);

    table.appendChild(tr);
  });
}

metricSel.addEventListener('change', render);
stationSel.addEventListener('change', render);

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
    print(f"HTML dashboard generated at {HTML_PATH}")

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
