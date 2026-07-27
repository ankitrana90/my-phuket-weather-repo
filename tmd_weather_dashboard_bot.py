"""
TMD Phuket Weather -> Raw Excel Log + Cinematic Apple Weather UI (Single Script).

Excel (Phuket_Weather.xlsx) stores RAW DATA ONLY. The generated HTML dashboard
(Phuket_Weather_Dashboard.html) features an Apple Weather / VisionOS inspired UI,
complete with particle sky backdrops, frosted liquid glass cards, interactive hourly
timeline tiles, and smooth Apple-style curved area charts.
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
<title>Phuket Weather Intelligence</title>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
  :root {
    --deep-navy: #102A54;
    --electric-blue: #2D7DFF;
    --azure: #4A90FF;
    --golden-white: #FFF4C6;
    --danger-orange: #FF8A3D;
    --glass-bg: rgba(255, 255, 255, 0.18);
    --glass-border: 1px solid rgba(255, 255, 255, 0.25);
    --glass-shadow: 0 25px 70px rgba(0, 0, 0, 0.15);
    --text-dark: #0B1320;
    --text-light: #FFFFFF;
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }
  
  body {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "SF Pro Text", sans-serif;
    background: var(--deep-navy);
    color: var(--text-light);
    min-height: 100vh;
    overflow-x: hidden;
    -webkit-font-smoothing: antialiased;
  }

  /* Sky Particle Backdrop Canvas */
  #skyCanvas {
    position: fixed;
    top: 0; left: 0;
    width: 100vw; height: 100vh;
    z-index: 1;
    pointer-events: none;
  }

  /* Layout Wrapper */
  .app-container {
    position: relative;
    z-index: 2;
    max-width: 1120px;
    margin: 0 auto;
    padding: 32px 24px 64px 24px;
  }

  /* Header Section */
  header {
    display: flex;
    justify-content: space-between;
    align-items: flex-end;
    margin-bottom: 24px;
  }

  .header-title h1 {
    font-size: 38px;
    font-weight: 300;
    letter-spacing: -0.8px;
  }

  .header-title .meta {
    font-size: 14px;
    color: rgba(255, 255, 255, 0.7);
    margin-top: 4px;
  }

  .station-picker {
    appearance: none;
    background: var(--glass-bg);
    border: var(--glass-border);
    backdrop-filter: blur(25px);
    color: var(--text-light);
    padding: 10px 18px;
    border-radius: 20px;
    font-size: 14px;
    font-weight: 400;
    outline: none;
    cursor: pointer;
    transition: transform 0.3s ease, background 0.3s ease;
  }
  .station-picker:hover {
    transform: translateY(-2px);
    background: rgba(255, 255, 255, 0.28);
  }
  .station-picker option {
    background: var(--deep-navy);
    color: var(--text-light);
  }

  /* Glass Card Base Styling */
  .glass-card {
    background: var(--glass-bg);
    border: var(--glass-border);
    backdrop-filter: blur(25px);
    border-radius: 28px;
    box-shadow: var(--glass-shadow);
    transition: transform 0.5s cubic-bezier(0.16, 1, 0.3, 1), background 0.3s ease;
  }
  .glass-card:hover {
    transform: translateY(-4px);
    background: rgba(255, 255, 255, 0.24);
  }

  /* Hero Display (Top 35% Screen Height) */
  .hero-section {
    padding: 36px 40px;
    margin-bottom: 28px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    position: relative;
    overflow: hidden;
  }

  .hero-main .condition-tag {
    display: inline-block;
    padding: 6px 14px;
    background: rgba(255, 255, 255, 0.2);
    border-radius: 14px;
    font-size: 13px;
    font-weight: 500;
    margin-bottom: 12px;
  }

  .hero-main .hero-temp {
    font-size: 64px;
    font-weight: 300;
    line-height: 1;
    letter-spacing: -2px;
  }

  .hero-metrics {
    display: flex;
    gap: 32px;
  }

  .hero-metric-item {
    text-align: right;
  }
  .hero-metric-item .lbl {
    font-size: 13px;
    color: rgba(255, 255, 255, 0.7);
    margin-bottom: 4px;
  }
  .hero-metric-item .val {
    font-size: 24px;
    font-weight: 400;
  }

  /* Current Conditions Cards Grid */
  .cards-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 20px;
    margin-bottom: 32px;
  }

  .condition-card {
    padding: 24px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    height: 180px;
  }

  .card-top {
    display: flex;
    justify-content: space-between;
    align-items: center;
  }

  .card-top .title {
    font-size: 14px;
    color: rgba(255, 255, 255, 0.75);
    font-weight: 400;
  }

  .card-icon {
    width: 22px;
    height: 22px;
    stroke: currentColor;
    stroke-width: 1.5;
    fill: none;
  }

  .card-val {
    font-size: 32px;
    font-weight: 300;
    letter-spacing: -1px;
    margin-top: 8px;
  }

  .card-sub {
    font-size: 12px;
    color: rgba(255, 255, 255, 0.65);
    margin-top: 4px;
  }

  /* Section Title */
  .section-title {
    font-size: 22px;
    font-weight: 300;
    letter-spacing: -0.5px;
    margin-bottom: 16px;
    display: flex;
    align-items: center;
    gap: 10px;
  }

  /* Timeline Strip */
  .timeline-container {
    display: flex;
    gap: 14px;
    overflow-x: auto;
    padding-bottom: 12px;
    margin-bottom: 36px;
    scrollbar-width: none;
  }
  .timeline-container::-webkit-scrollbar { display: none; }

  .time-tile {
    min-width: 110px;
    padding: 20px 16px;
    text-align: center;
    flex-shrink: 0;
  }

  .time-tile .t-hour {
    font-size: 14px;
    color: rgba(255, 255, 255, 0.7);
    margin-bottom: 12px;
  }

  .time-tile .t-val {
    font-size: 18px;
    font-weight: 400;
    margin: 8px 0;
  }

  .time-tile .t-sub {
    font-size: 11px;
    color: rgba(255, 255, 255, 0.6);
  }

  /* Day Accordion Cards Section */
  .accordion-section {
    display: flex;
    flex-direction: column;
    gap: 16px;
  }

  .day-accordion {
    padding: 24px 28px;
    cursor: pointer;
  }

  .day-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
  }

  .day-info {
    display: flex;
    align-items: center;
    gap: 24px;
  }

  .day-name {
    font-size: 18px;
    font-weight: 400;
    width: 120px;
  }

  .day-summary-text {
    font-size: 14px;
    color: rgba(255, 255, 255, 0.75);
  }

  .day-ring-container {
    display: flex;
    align-items: center;
    gap: 16px;
  }

  /* Circular Liquid Meter */
  .ring-meter {
    position: relative;
    width: 48px;
    height: 48px;
  }

  .ring-meter svg {
    transform: rotate(-90deg);
    width: 48px;
    height: 48px;
  }

  .ring-meter circle {
    fill: none;
    stroke-width: 4;
  }

  .ring-bg { stroke: rgba(255, 255, 255, 0.15); }
  .ring-fill {
    stroke: var(--electric-blue);
    stroke-dasharray: 126;
    stroke-dashoffset: 126;
    transition: stroke-dashoffset 1s ease;
  }

  .ring-val {
    position: absolute;
    inset: 0;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 10px;
    font-weight: 500;
  }

  .day-body {
    max-height: 0;
    overflow: hidden;
    transition: max-height 0.6s cubic-bezier(0.16, 1, 0.3, 1), margin-top 0.4s ease;
  }

  .day-accordion.open .day-body {
    max-height: 240px;
    margin-top: 24px;
  }

  .chart-canvas {
    width: 100%;
    height: 160px;
  }

  /* Responsive Breakpoints */
  @media (max-width: 900px) {
    .cards-grid { grid-template-columns: repeat(2, 1fr); }
    .hero-section { flex-direction: column; align-items: flex-start; gap: 24px; }
    .hero-metrics { text-align: left; }
    .hero-metric-item { text-align: left; }
  }
  @media (max-width: 550px) {
    .cards-grid { grid-template-columns: 1fr; }
  }
</style>
</head>
<body>

  <!-- Sky Canvas Backdrop -->
  <canvas id="skyCanvas"></canvas>

  <div class="app-container">
    <header>
      <div class="header-title">
        <h1>Phuket, Thailand</h1>
        <div class="meta" id="headerMeta">Syncing telemetry...</div>
      </div>
      <select id="stationPicker" class="station-picker"></select>
    </header>

    <!-- HERO DISPLAY -->
    <div class="glass-card hero-section">
      <div class="hero-main">
        <div class="condition-tag" id="heroConditionTag">Atmospheric Scan</div>
        <div class="hero-temp" id="heroVal">-- mm</div>
      </div>
      <div class="hero-metrics">
        <div class="hero-metric-item">
          <div class="lbl">Wind Flow</div>
          <div class="val" id="heroWind">-- km/h</div>
        </div>
        <div class="hero-metric-item">
          <div class="lbl">Visibility</div>
          <div class="val" id="heroVis">-- km</div>
        </div>
        <div class="hero-metric-item">
          <div class="lbl">Thunder Risk</div>
          <div class="val" id="heroThunder">Low</div>
        </div>
      </div>
    </div>

    <!-- CURRENT CONDITIONS CARDS GRID -->
    <div class="cards-grid">
      <!-- Rain Card -->
      <div class="glass-card condition-card">
        <div class="card-top">
          <span class="title">Precipitation</span>
          <svg class="card-icon" viewBox="0 0 24 24"><path d="M20 16.58A5 5 0 0 0 18 7h-1.26A8 8 0 1 0 4 15.25"/><path d="M8 16v4M12 18v4M16 16v4"/></svg>
        </div>
        <div>
          <div class="card-val" id="cardRain">-- mm</div>
          <div class="card-sub" id="cardRainSub">Current Reading</div>
        </div>
      </div>

      <!-- Wind Card -->
      <div class="glass-card condition-card">
        <div class="card-top">
          <span class="title">Wind Speed</span>
          <svg class="card-icon" viewBox="0 0 24 24"><path d="M17.7 7.7a2.5 2.5 0 1 1 1.8 4.3H2M9.6 4.6A2 2 0 1 1 11 8H2M12.6 19.4A2 2 0 1 0 14 16H2"/></svg>
        </div>
        <div>
          <div class="card-val" id="cardWind">-- km/h</div>
          <div class="card-sub">Air Velocity</div>
        </div>
      </div>

      <!-- Visibility Card -->
      <div class="glass-card condition-card">
        <div class="card-top">
          <span class="title">Visibility</span>
          <svg class="card-icon" viewBox="0 0 24 24"><path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7Z"/><circle cx="12" cy="12" r="3"/></svg>
        </div>
        <div>
          <div class="card-val" id="cardVis">-- km</div>
          <div class="card-sub" id="cardVisSub">Clarity Depth</div>
        </div>
      </div>

      <!-- Thunder Risk Card -->
      <div class="glass-card condition-card">
        <div class="card-top">
          <span class="title">Thunder Risk</span>
          <svg class="card-icon" viewBox="0 0 24 24"><path d="M13 2 3 14h9l-1 8 10-12h-9l1-8z"/></svg>
        </div>
        <div>
          <div class="card-val" id="cardThunder">Minimal</div>
          <div class="card-sub">Electrical Activity</div>
        </div>
      </div>
    </div>

    <!-- FORECAST TIMELINE -->
    <div class="section-title">Hourly Forecast</div>
    <div class="timeline-container" id="timelineStrip"></div>

    <!-- ACCORDION DAY CARDS -->
    <div class="section-title">Historical & Daily Trends</div>
    <div class="accordion-section" id="accordionSection"></div>
  </div>

<script>
const CONFIG = __CONFIG_JSON__;
const stationPicker = document.getElementById('stationPicker');
let DATA = null;

// Populate Station Options
CONFIG.stations.forEach(s => {
  const opt = document.createElement('option');
  opt.value = s.code; opt.textContent = s.label_display;
  if (s.code === 'COMBINED') opt.selected = true;
  stationPicker.appendChild(opt);
});

// Auto-Fetch live dataset from repository
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
    
    document.getElementById('headerMeta').textContent = `Updated Live • ${new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}`;
    renderApp();
  } catch (err) {
    document.getElementById('headerMeta').textContent = 'Displaying static atmospheric cache';
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

/* ------------------------------------------------------------------
   PARTICLE BACKDROP CANVAS (Sky Engine)
------------------------------------------------------------------ */
let skyAnimationId;
function initSkyEngine(rainAmount) {
  const canvas = document.getElementById('skyCanvas');
  const ctx = canvas.getContext('2d');
  canvas.width = window.innerWidth;
  canvas.height = window.innerHeight;

  cancelAnimationFrame(skyAnimationId);

  const particles = [];
  const count = Math.min(Math.floor(rainAmount * 30) + 20, 300);

  for (let i = 0; i < count; i++) {
    particles.push({
      x: Math.random() * canvas.width,
      y: Math.random() * canvas.height,
      len: Math.random() * 20 + 10,
      speed: Math.random() * 10 + 6 + (rainAmount > 10 ? 8 : 0)
    });
  }

  let flash = 0;

  function renderSky() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    // Sky Background Gradient
    const bgGrad = ctx.createLinearGradient(0, 0, 0, canvas.height);
    if (rainAmount > 15) {
      bgGrad.addColorStop(0, '#0F172A');
      bgGrad.addColorStop(1, '#1E293B');
    } else if (rainAmount > 2) {
      bgGrad.addColorStop(0, '#102A54');
      bgGrad.addColorStop(1, '#1E3A8A');
    } else {
      bgGrad.addColorStop(0, '#102A54');
      bgGrad.addColorStop(1, '#1D4ED8');
    }
    ctx.fillStyle = bgGrad;
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    // Lightning Trigger for Severe Rain (>15mm)
    if (rainAmount > 15 && Math.random() < 0.006) flash = 0.8;
    if (flash > 0) {
      ctx.fillStyle = `rgba(255, 244, 198, ${flash})`;
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      flash -= 0.04;
    }

    // Rain Particle Rendering
    if (rainAmount > 0) {
      ctx.strokeStyle = 'rgba(74, 144, 255, 0.45)';
      ctx.lineWidth = 1.2;
      ctx.beginPath();
      particles.forEach(p => {
        ctx.moveTo(p.x, p.y);
        ctx.lineTo(p.x - 2, p.y + p.len);
        p.y += p.speed;
        p.x -= 0.8;
        if (p.y > canvas.height) {
          p.y = -p.len;
          p.x = Math.random() * canvas.width;
        }
      });
      ctx.stroke();
    }

    skyAnimationId = requestAnimationFrame(renderSky);
  }
  renderSky();
}

/* ------------------------------------------------------------------
   MAIN APP RENDERER
------------------------------------------------------------------ */
function renderApp() {
  if (!DATA) return;

  const station = stationPicker.value;
  const rainData = DATA.payload.Rain[station];
  const windData = DATA.payload.Wind[station];
  const visData = DATA.payload.Vis[station];

  const latestRain = rainData.grid.flat().filter(v => v !== null).pop() || 0;
  const latestWind = windData.grid.flat().filter(v => v !== null).pop() || 0;
  const latestVis = visData.grid.flat().filter(v => v !== null).pop() || 10;

  // Hero Display Updates
  document.getElementById('heroVal').textContent = `${latestRain} mm`;
  document.getElementById('heroWind').textContent = `${latestWind} km/h`;
  document.getElementById('heroVis').textContent = `${latestVis} km`;
  
  const thunderRisk = latestRain > 15 ? 'High' : latestRain > 5 ? 'Moderate' : 'Low';
  document.getElementById('heroThunder').textContent = thunderRisk;

  document.getElementById('heroConditionTag').textContent = 
    latestRain > 15 ? "Severe Monsoon Stream" :
    latestRain > 2  ? "Active Rainfall" : "Clear Sky State";

  // Cards Grid Updates
  document.getElementById('cardRain').textContent = `${latestRain} mm`;
  document.getElementById('cardWind').textContent = `${latestWind} km/h`;
  document.getElementById('cardVis').textContent = `${latestVis} km`;
  document.getElementById('cardThunder').textContent = thunderRisk;

  document.getElementById('cardVisSub').textContent = 
    latestVis < 5 ? "Hazy Atmosphere" : "Clear Horizon";

  initSkyEngine(latestRain);

  // Render Forecast Timeline Tiles
  const timelineStrip = document.getElementById('timelineStrip');
  timelineStrip.innerHTML = '';

  const lastDayIndex = DATA.dates.length - 1;
  const hourlyTimes = DATA.times;
  const hourlyRain = rainData.grid[lastDayIndex] || [];
  const hourlyWind = windData.grid[lastDayIndex] || [];

  hourlyTimes.forEach((t, idx) => {
    const rVal = hourlyRain[idx] !== null ? hourlyRain[idx] : 0;
    const wVal = hourlyWind[idx] !== null ? hourlyWind[idx] : 0;

    const tile = document.createElement('div');
    tile.className = 'glass-card time-tile';
    tile.innerHTML = `
      <div class="t-hour">${t}</div>
      <svg class="card-icon" style="margin: 0 auto;" viewBox="0 0 24 24">
        <path d="M20 16.58A5 5 0 0 0 18 7h-1.26A8 8 0 1 0 4 15.25"/><path d="M8 16v4M12 18v4M16 16v4"/>
      </svg>
      <div class="t-val">${rVal} mm</div>
      <div class="t-sub">${wVal} km/h</div>
    `;
    timelineStrip.appendChild(tile);
  });

  // Render Accordion Day Cards
  const accordionSection = document.getElementById('accordionSection');
  accordionSection.innerHTML = '';

  DATA.dates.forEach((d, i) => {
    const totalRain = rainData.summary[i] !== null ? rainData.summary[i] : 0;
    const fillPercent = Math.min(100, (totalRain / CONFIG.rainfall_cap) * 100);
    const strokeOffset = 126 - (126 * fillPercent) / 100;

    const card = document.createElement('div');
    card.className = 'glass-card day-accordion';
    card.innerHTML = `
      <div class="day-header">
        <div class="day-info">
          <div class="day-name">${d}</div>
          <div class="day-summary-text">${totalRain > 10 ? 'Heavy Precipitation' : totalRain > 0 ? 'Light Drizzle' : 'Clear Day'}</div>
        </div>
        <div class="day-ring-container">
          <div class="ring-meter">
            <svg>
              <circle class="ring-bg" cx="24" cy="24" r="20"/>
              <circle class="ring-fill" cx="24" cy="24" r="20" style="stroke-dashoffset: ${strokeOffset};"/>
            </svg>
            <div class="ring-val">${totalRain}</div>
          </div>
        </div>
      </div>
      <div class="day-body">
        <canvas class="chart-canvas" id="chart-${i}"></canvas>
      </div>
    `;

    card.addEventListener('click', () => {
      card.classList.toggle('open');
      if (card.classList.contains('open')) {
        renderDayChart(`chart-${i}`, DATA.times, rainData.grid[i]);
      }
    });

    accordionSection.appendChild(card);
  });
}

/* Apple-Style Smooth Curved Area Chart */
function renderDayChart(canvasId, times, values) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  
  canvas.width = canvas.offsetWidth;
  canvas.height = canvas.offsetHeight;

  const padding = 24;
  const w = canvas.width - padding * 2;
  const h = canvas.height - padding * 2;

  const validVals = values.filter(v => v !== null);
  const maxVal = Math.max(...validVals, 8);

  const points = values.map((v, idx) => {
    const val = v !== null ? v : 0;
    const x = padding + (idx / (values.length - 1)) * w;
    const y = canvas.height - padding - (val / maxVal) * h;
    return { x, y };
  });

  ctx.clearRect(0, 0, canvas.width, canvas.height);

  // Gradient Fill
  const fillGrad = ctx.createLinearGradient(0, 0, 0, canvas.height);
  fillGrad.addColorStop(0, 'rgba(45, 125, 255, 0.4)');
  fillGrad.addColorStop(1, 'rgba(45, 125, 255, 0.0)');

  ctx.beginPath();
  ctx.moveTo(points[0].x, points[0].y);
  for (let i = 0; i < points.length - 1; i++) {
    const xc = (points[i].x + points[i + 1].x) / 2;
    const yc = (points[i].y + points[i + 1].y) / 2;
    ctx.quadraticCurveTo(points[i].x, points[i].y, xc, yc);
  }
  ctx.lineTo(points[points.length - 1].x, points[points.length - 1].y);
  ctx.lineTo(points[points.length - 1].x, canvas.height - padding);
  ctx.lineTo(points[0].x, canvas.height - padding);
  ctx.closePath();
  ctx.fillStyle = fillGrad;
  ctx.fill();

  // Curve Line
  ctx.beginPath();
  ctx.moveTo(points[0].x, points[0].y);
  for (let i = 0; i < points.length - 1; i++) {
    const xc = (points[i].x + points[i + 1].x) / 2;
    const yc = (points[i].y + points[i + 1].y) / 2;
    ctx.quadraticCurveTo(points[i].x, points[i].y, xc, yc);
  }
  ctx.strokeStyle = '#2D7DFF';
  ctx.lineWidth = 2.5;
  ctx.stroke();
}

stationPicker.addEventListener('change', renderApp);
window.addEventListener('resize', renderApp);
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
    print(f"Apple Weather UI HTML generated at {HTML_PATH}")

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
