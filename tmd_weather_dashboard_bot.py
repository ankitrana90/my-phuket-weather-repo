"""
TMD Phuket Weather -> Raw Excel Log + Advanced Physics Apple Weather Matrix (Single Script).

Excel (Phuket_Weather.xlsx) logs raw TMD observations alongside real-time condition descriptions
fetched from Open-Meteo API. The generated HTML dashboard (Phuket_Weather_Dashboard.html) features:
  - Dates in DESCENDING order (newest on top).
  - Clean cells (zeros hidden).
  - 8-Slot real-time condition emoji row beneath each Date header mapped from Excel.
  - Hero display synced with the latest recorded weather condition emoji from Excel.
  - Physics-based fluid dynamics (gravity wave propagation, sloshing, and splash mechanics).
  - Color gradient scaling from light blue to dark navy at >= 8mm.
  - Rain particles tilted dynamically up to 60 degrees based on wind speed.
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

# Open-Meteo Free Weather API for Phuket (Lat: 7.8804, Lon: 98.3923)
OPEN_METEO_URL = "https://api.open-meteo.com/v1/forecast?latitude=7.8804&longitude=98.3923&current_weather=true"

TARGET_STATION_NAMES = {"PHUKET AIRPORT", "PHUKET"}
TARGET_WMO_NUMBERS = {"48565", "48564"}

VISIBILITY_TAG_CANDIDATES = ["VisibilityLand", "LandVisibility", "Visibility", "Vis"]

OUTPUT_DIR = Path(".")
EXCEL_PATH = OUTPUT_DIR / "Phuket_Weather.xlsx"
HTML_PATH = OUTPUT_DIR / "Phuket_Weather_Dashboard.html"

DATA_SHEET = "Phuket Weather"

LEGACY_SHEET_NAMES_TO_DROP = {"Dashboard", "Lists"}
LEGACY_SHEET_PREFIXES_TO_DROP = ("_pivot_",)

HEADERS = ["Station", "WmoStationNumber", "DateTime", "WindSpeed", "Rainfall_mm", "Rainfall24Hr_mm", "LandVisibility", "Condition"]

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

# WMO Weather Interpretation Codes (Open-Meteo)
WMO_CODE_MAP = {
    0: ("Clear sky", "☀️"),
    1: ("Mainly clear", "🌤️"),
    2: ("Partly cloudy", "⛅"),
    3: ("Overcast", "☁️"),
    45: ("Fog", "🌫️"),
    48: ("Depositing rime fog", "🌫️"),
    51: ("Light drizzle", "🌦️"),
    53: ("Moderate drizzle", "🌦️"),
    55: ("Dense drizzle", "🌧️"),
    61: ("Slight rain", "🌦️"),
    63: ("Moderate rain", "🌧️"),
    65: ("Heavy rain", "🌧️"),
    80: ("Slight rain showers", "🌦️"),
    81: ("Moderate rain showers", "🌧️"),
    82: ("Violent rain showers", "⛈️"),
    95: ("Thunderstorm", "⛈️"),
    96: ("Thunderstorm with slight hail", "⛈️"),
    99: ("Thunderstorm with heavy hail", "⛈️"),
}

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

def fetch_live_weather_condition():
    """Fetches real-time condition description and emoji from Open-Meteo API."""
    try:
        resp = requests.get(OPEN_METEO_URL, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        current = data.get("current_weather", {})
        code = int(current.get("weathercode", 0))
        text, emoji = WMO_CODE_MAP.get(code, ("Variable Conditions", "🌤️"))
        return f"{emoji} {text}"
    except Exception as e:
        print(f"Warning: Could not fetch Open-Meteo condition ({e}). Using fallback.")
        return "🌤️ Variable Conditions"

# ----------------------------------------------------------------------------
# Fetch & Parse
# ----------------------------------------------------------------------------

def fetch_raw_records():
    live_condition = fetch_live_weather_condition()
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
            "condition": live_condition
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
        cond = recs[0]["condition"] if recs else "🌤️ Variable Conditions"

        for r in recs:
            rows.append({
                "station": r["station"],
                "wmo": r["wmo"],
                "datetime": dt,
                "wind_speed": none_if_nan(r["wind_speed"]),
                "rainfall_mm": none_if_nan(r["rainfall_mm"]),
                "rainfall_24hr_mm": none_if_nan(r["rainfall_24hr_mm"]),
                "land_visibility": none_if_nan(r["land_visibility"]),
                "condition": r["condition"]
            })

        rows.append({
            "station": STATION_LABELS["COMBINED"],
            "wmo": "+".join(wmo_list),
            "datetime": dt,
            "wind_speed": mean_ignore_nan([r["wind_speed"] for r in recs]),
            "rainfall_mm": mean_ignore_nan([r["rainfall_mm"] for r in recs]),
            "rainfall_24hr_mm": mean_ignore_nan([r["rainfall_24hr_mm"] for r in recs]),
            "land_visibility": mean_ignore_nan([r["land_visibility"] for r in recs]),
            "condition": cond
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
        ws.append([r["station"], r["wmo"], r["datetime"], r["wind_speed"], r["rainfall_mm"], r["rainfall_24hr_mm"], r["land_visibility"], r.get("condition", "")])
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
<title>Phuket Weather Intelligence</title>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
  :root {
    --bg-light: #F8FAFC;
    --card-bg: rgba(255, 255, 255, 0.85);
    --card-border: 1px solid rgba(226, 232, 240, 0.8);
    --card-shadow: 0 15px 35px -10px rgba(15, 23, 42, 0.05);
    --text-primary: #0F172A;
    --text-secondary: #64748B;
    --accent-blue: #3B82F6;
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "SF Pro Text", sans-serif;
    background-color: var(--bg-light);
    color: var(--text-primary);
    min-height: 100vh;
    padding: 32px 24px;
    -webkit-font-smoothing: antialiased;
  }

  .container {
    max-width: 1280px;
    margin: 0 auto;
  }

  header {
    display: flex;
    justify-content: space-between;
    align-items: flex-end;
    margin-bottom: 24px;
  }

  .location-info h1 {
    font-size: 32px;
    font-weight: 600;
    letter-spacing: -0.8px;
    color: var(--text-primary);
  }

  .location-info .meta {
    font-size: 14px;
    color: var(--text-secondary);
    margin-top: 4px;
  }

  .station-select {
    appearance: none;
    background: var(--card-bg);
    border: var(--card-border);
    backdrop-filter: blur(20px);
    color: var(--text-primary);
    padding: 10px 18px;
    border-radius: 18px;
    font-size: 14px;
    font-weight: 500;
    outline: none;
    cursor: pointer;
    box-shadow: var(--card-shadow);
    transition: all 0.2s ease;
  }
  .station-select:hover {
    transform: translateY(-2px);
    border-color: var(--accent-blue);
  }

  /* Hero Display Card */
  .hero-card {
    background: var(--card-bg);
    border: var(--card-border);
    backdrop-filter: blur(25px);
    border-radius: 28px;
    padding: 28px 32px;
    box-shadow: var(--card-shadow);
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 28px;
  }

  .hero-left {
    display: flex;
    align-items: center;
    gap: 20px;
  }

  .hero-emoji {
    font-size: 48px;
    line-height: 1;
  }

  .hero-val {
    font-size: 44px;
    font-weight: 300;
    letter-spacing: -1.5px;
    line-height: 1;
  }

  .hero-status {
    font-size: 14px;
    color: var(--text-secondary);
    margin-top: 6px;
  }

  .grid-title {
    font-size: 20px;
    font-weight: 600;
    letter-spacing: -0.4px;
    margin-bottom: 16px;
  }

  .matrix-wrapper {
    background: var(--card-bg);
    border: var(--card-border);
    backdrop-filter: blur(25px);
    border-radius: 24px;
    box-shadow: var(--card-shadow);
    overflow: auto;
    padding: 12px;
  }

  table {
    width: 100%;
    border-collapse: separate;
    border-spacing: 10px;
    font-size: 13px;
    text-align: center;
  }

  th {
    padding: 12px 8px;
    color: var(--text-secondary);
    font-weight: 600;
    font-size: 13px;
    letter-spacing: -0.2px;
  }

  th.date-col { text-align: left; padding-left: 16px; min-width: 175px; }
  th.summary-col { min-width: 240px; text-align: left; padding-left: 16px; }

  td.date-cell {
    text-align: left;
    padding: 12px 14px;
    background: rgba(255, 255, 255, 0.6);
    border-radius: 16px;
    border: 1px solid rgba(226, 232, 240, 0.6);
  }

  .date-title { font-size: 15px; font-weight: 600; color: var(--text-primary); }

  /* 8-Slot Time Emoji Strip Under Date */
  .emoji-slots-strip {
    display: flex;
    gap: 3px;
    margin-top: 8px;
    align-items: center;
  }

  .slot-emoji {
    font-size: 12px;
    width: 16px;
    height: 18px;
    display: flex;
    align-items: center;
    justify-content: center;
  }

  .slot-emoji.empty-emoji {
    opacity: 0;
  }

  /* Interactive Physics Fluid Cell */
  td.time-cell {
    position: relative;
    width: 82px;
    height: 80px;
    background: rgba(255, 255, 255, 0.6);
    border-radius: 16px;
    border: 1px solid rgba(226, 232, 240, 0.6);
    overflow: hidden;
    vertical-align: middle;
    transition: transform 0.2s ease, border-color 0.2s ease;
  }

  td.time-cell:hover {
    transform: scale(1.05);
    border-color: var(--accent-blue);
  }

  .cell-canvas {
    position: absolute;
    inset: 0;
    width: 100%; height: 100%;
    pointer-events: none;
    z-index: 1;
  }

  .cell-content {
    position: relative;
    z-index: 3;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    height: 100%;
    pointer-events: none;
  }

  .val-num {
    font-size: 15px;
    font-weight: 700;
    letter-spacing: -0.3px;
    text-shadow: 0 1px 2px rgba(255, 255, 255, 0.8);
  }

  .val-unit { font-size: 10px; color: var(--text-secondary); font-weight: 500; }

  td.summary-cell {
    text-align: left;
    padding: 12px 18px;
    background: rgba(59, 130, 246, 0.06);
    border-radius: 16px;
    border: 1px solid rgba(59, 130, 246, 0.15);
  }

  .summary-val { font-size: 17px; font-weight: 600; color: var(--accent-blue); }
  .summary-commentary { font-size: 12px; color: var(--text-secondary); line-height: 1.35; margin-top: 4px; }

  @media (max-width: 768px) {
    .hero-card { flex-direction: column; align-items: flex-start; gap: 16px; }
  }
</style>
</head>
<body>

  <div class="container">
    <header>
      <div class="location-info">
        <h1>Phuket, Thailand</h1>
        <div class="meta" id="headerMeta">Syncing live telemetry...</div>
      </div>
      <select id="stationSelect" class="station-select"></select>
    </header>

    <!-- Hero Display Card -->
    <div class="hero-card">
      <div class="hero-left">
        <div class="hero-emoji" id="heroEmoji">🌤️</div>
        <div>
          <div class="hero-val" id="heroVal">-- mm</div>
          <div class="hero-status" id="heroStatus">Scanning precipitation pattern...</div>
        </div>
      </div>
    </div>

    <div class="grid-title">Time-of-Day Rainfall Matrix (Descending Date Order)</div>
    
    <div class="matrix-wrapper">
      <table id="matrixTable"></table>
    </div>
  </div>

<script>
const CONFIG = __CONFIG_JSON__;
const stationSelect = document.getElementById('stationSelect');
let DATA = null;
let activeCanvasRenderers = [];

CONFIG.stations.forEach(s => {
  const opt = document.createElement('option');
  opt.value = s.code; opt.textContent = s.label_display;
  if (s.code === 'COMBINED') opt.selected = true;
  stationSelect.appendChild(opt);
});

async function autoFetchServerData() {
  try {
    const res = await fetch(CONFIG.xlsx_filename, { cache: 'no-cache' });
    if (!resThis usually comes down to two distinct issues: **formatting/aggregation handling** between Excel and your dashboard, and **evaluation context/sorting** for the latest record.

Here are the most common root causes and how to fix them:

---

## 1. Why the Emoji is Missing Near the Date in the Dashboard

* **Measure vs. Column Aggregation:** If your dashboard treats the date or status as a measure/metric rather than a dimension/row, it might be trying to aggregate the text (e.g., taking `FIRST()` or `MAX()`) and failing to display the emoji.
  * **Fix:** Ensure the field containing the emoji is set as an **unaggregated dimension/attribute** or categorized as plain text.
* **Date Format Mismatch in Relationships:** If the date format in Excel (e.g., `DD/MM/YYYY`) is parsed differently by the dashboard engine (e.g., `MM/DD/YYYY` or string), the conditional logic fails to find a match and returns `BLANK` or `NULL`.
  * **Fix:** Convert both date columns explicitly to a standard **Date/Time data type** inside your dashboard's data model or query editor before applying conditional logic.
* **Unicode/Font Rendering:** Some dashboard server environments (especially web renderers) strip UTF-8 emojis or default to system fonts that don't support color emojis.
  * **Fix:** Try replacing high-range emojis with standard web-safe icons, SVG indicators, or basic Unicode symbols (e.g., `▲`, `▼`, `●`) as a test.

---

## 2. Why the Top Emoji (Latest Record) Is Showing the Wrong Condition

* **Text-Based Date Sorting:** If the date column is stored as text, sorting evaluates alphabetically (e.g., `"15/01/2026"` comes after `"01/08/2026"`), picking the wrong "latest" row.
  * **Fix:** Sort explicitly by a raw `DateTime` column or integer timestamp (`YYYYMMDD`).
* **Condition Evaluated on Summary instead of Row:** A common DAX/SQL mistake is evaluating the condition on an aggregated metric (e.g., `AVERAGE` or `SUM`) across the entire table rather than filtering to the single row corresponding to `MAX(Date)`.
  * **Fix Structure (Logic Flow):**
    1. Identify `Max_Date = MAX(Table[Date])`
    2. Filter table to `Table[Date] = Max_Date`
    3. Retrieve the status/value **for that specific row**
    4. Apply the `IF` / `SWITCH` emoji logic to *that* single value.

---

<ElicitationsGroup message="To give you the exact formula or setup fix:">
  <Elicitation label="Share dashboard platform and formula logic" query="I am using [Power BI / Tableau / Excel / Custom App]. Here is the formula or logic I'm using to show the emoji:" />
</ElicitationsGroup>
