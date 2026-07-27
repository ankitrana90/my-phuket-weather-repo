"""
TMD Phuket Weather -> Raw Excel Log + Apple Weather Atmospheric Dashboard (Single Script).

Excel (Phuket_Weather.xlsx) stores RAW DATA ONLY. The interactive HTML dashboard
(Phuket_Weather_Dashboard.html) features an Apple Weather-inspired atmospheric UI,
complete with particle fluid streamlines, atmospheric rain simulations, depth-based fog,
and interactive trend analytics.
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
    --card-bg: rgba(255, 255, 255, 0.72);
    --card-border: rgba(232, 237, 243, 0.8);
    --glass-shadow: 0 20px 40px -15px rgba(16, 20, 24, 0.05);
    --primary-blue: #3A6FF7;
    --cloud-gray: #E8EDF3;
    --text-primary: #101418;
    --text-muted: #64748B;
    --accent-white: #FFFFFF;
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "SF Pro Text", "Helvetica Neue", sans-serif;
    background-color: var(--bg-color);
    color: var(--text-primary);
    padding: 32px 24px;
    max-width: 1200px;
    margin: 0 auto;
    -webkit-font-smoothing: antialiased;
  }

  /* Header Section */
  header {
    display: flex;
    justify-content: space-between;
    align-items: flex-end;
    margin-bottom: 28px;
  }

  .location-title {
    font-size: 32px;
    font-weight: 600;
    letter-spacing: -0.8px;
    color: var(--text-primary);
  }

  .subtitle {
    font-size: 14px;
    color: var(--text-muted);
    font-weight: 400;
    margin-top: 4px;
  }

  .station-select {
    appearance: none;
    background: var(--card-bg);
    border: 1px solid var(--card-border);
    padding: 10px 16px;
    border-radius: 20px;
    font-size: 13px;
    font-weight: 500;
    color: var(--text-primary);
    backdrop-filter: blur(20px);
    cursor: pointer;
    box-shadow: var(--glass-shadow);
    outline: none;
    transition: all 0.2s ease;
  }
  .station-select:hover { border-color: var(--primary-blue); }

  /* Grid Layout */
  .dashboard-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 20px;
  }

  /* Shared Card Styling */
  .card {
    background: var(--card-bg);
    border: 1px solid var(--card-border);
    border-radius: 24px;
    backdrop-filter: blur(24px);
    box-shadow: var(--glass-shadow);
    overflow: hidden;
    position: relative;
  }

  /* Hero Card: Full Width Rainfall */
  .hero-card {
    grid-column: span 2;
    height: 380px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
  }

  .hero-canvas {
    position: absolute;
    inset: 0;
    width: 100%;
    height: 100%;
    z-index: 1;
  }

  .card-overlay {
    position: relative;
    z-index: 2;
    padding: 28px;
    pointer-events: none;
  }

  .floating-glass-pill {
    display: inline-block;
    background: rgba(255, 255, 255, 0.65);
    border: 1px solid rgba(255, 255, 255, 0.8);
    backdrop-filter: blur(16px);
    padding: 16px 24px;
    border-radius: 20px;
    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.04);
  }

  .metric-label {
    font-size: 12px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    color: var(--text-muted);
  }

  .metric-value {
    font-size: 42px;
    font-weight: 600;
    letter-spacing: -1.2px;
    color: var(--text-primary);
    margin: 2px 0;
  }

  .metric-status {
    font-size: 13px;
    font-weight: 500;
    color: var(--primary-blue);
  }

  /* Secondary Cards */
  .secondary-card {
    height: 280px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
  }

  /* Trend Analytics Section */
  .trends-section {
    grid-column: span 2;
    margin-top: 12px;
    padding: 28px;
  }

  .chart-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 20px;
  }

  .chart-container {
    width: 100%;
    height: 160px;
    position: relative;
  }

  canvas.chart-canvas {
    width: 100%;
    height: 100%;
  }

  /* Interactive Sparkline Tooltip */
  .chart-tooltip {
    position: absolute;
    display: none;
    background: rgba(16, 20, 24, 0.85);
    color: #FFF;
    padding: 6px 12px;
    border-radius: 8px;
    font-size: 11px;
    font-weight: 500;
    pointer-events: none;
    transform: translate(-50%, -120%);
  }

  /* Responsive Adjustments */
  @media (max-width: 768px) {
    .dashboard-grid { grid-template-columns: 1fr; }
    .hero-card { grid-column: span 1; height: 320px; }
    .trends-section { grid-column: span 1; }
  }
</style>
</head>
<body>

  <header>
    <div>
      <h1 class="location-title">Phuket, Thailand</h1>
      <div class="subtitle" id="timestamp">Syncing live atmospheric telemetry...</div>
    </div>
    <select id="stationSelect" class="station-select"></select>
  </header>

  <div class="dashboard-grid">

    <!-- HERO CARD: RAINFALL EXPERIENCE -->
    <div class="card hero-card">
      <canvas id="rainCanvas" class="hero-canvas"></canvas>
      <div class="card-overlay">
        <div class="floating-glass-pill">
          <div class="metric-label">Precipitation</div>
          <div class="metric-value" id="rainVal">-- <span style="font-size:20px">mm</span></div>
          <div class="metric-status" id="rainStatus">Calm Atmospheric State</div>
        </div>
      </div>
    </div>

    <!-- WIND STREAMLINES CARD -->
    <div class="card secondary-card">
      <canvas id="windCanvas" class="hero-canvas"></canvas>
      <div class="card-overlay">
        <div class="floating-glass-pill">
          <div class="metric-label">Wind Velocity</div>
          <div class="metric-value" id="windVal">-- <span style="font-size:20px">km/h</span></div>
          <div class="metric-status" id="windStatus">Directional Flow</div>
        </div>
      </div>
    </div>

    <!-- VISIBILITY HORIZON CARD -->
    <div class="card secondary-card">
      <canvas id="visCanvas" class="hero-canvas"></canvas>
      <div class="card-overlay">
        <div class="floating-glass-pill">
          <div class="metric-label">Visual Depth</div>
          <div class="metric-value" id="visVal">-- <span style="font-size:20px">km</span></div>
          <div class="metric-status" id="visStatus">Atmospheric Clarity</div>
        </div>
      </div>
    </div>

    <!-- HISTORICAL TREND ANALYTICS -->
    <div class="card trends-section">
      <div class="chart-header">
        <div>
          <div class="metric-label">Atmospheric Trends</div>
          <div style="font-size: 18px; font-weight:600; margin-top:2px;">Historical Records</div>
        </div>
      </div>
      <div class="chart-container" id="chartContainer">
        <div class="chart-tooltip" id="chartTooltip"></div>
        <canvas id="trendChart" class="chart-canvas"></canvas>
      </div>
    </div>

  </div>

<script>
const CONFIG = __CONFIG_JSON__;
const stationSelect = document.getElementById('stationSelect');
let DATA = null;

// Populate Station Dropdown
CONFIG.stations.forEach(s => {
  const opt = document.createElement('option');
  opt.value = s.code; opt.textContent = s.label_display;
  if (s.code === 'COMBINED') opt.selected = true;
  stationSelect.appendChild(opt);
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
    
    document.getElementById('timestamp').textContent = `Updated Live • ${new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}`;
    renderDashboard();
  } catch (err) {
    document.getElementById('timestamp').textContent = 'Displaying static atmospheric telemetry cache';
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
      payload[m.code][s.code] = { grid, dates, times };
    });
  });

  return payload;
}

/* ------------------------------------------------------------------
   ATMOSPHERIC CANVASES (Rain, Wind, Visibility)
------------------------------------------------------------------ */

// 1. Rain Atmospheric Experience
let rainAnimationId;
function initRainCanvas(rainfall) {
  const canvas = document.getElementById('rainCanvas');
  const ctx = canvas.getContext('2d');
  canvas.width = canvas.offsetWidth;
  canvas.height = canvas.offsetHeight;

  cancelAnimationFrame(rainAnimationId);

  const drops = [];
  const count = Math.min(Math.floor(rainfall * 25) + 15, 350);
  
  for (let i = 0; i < count; i++) {
    drops.push({
      x: Math.random() * canvas.width,
      y: Math.random() * canvas.height,
      length: Math.random() * 18 + 10,
      speed: Math.random() * 8 + (rainfall > 10 ? 12 : 4),
      opacity: Math.random() * 0.4 + 0.2
    });
  }

  let flash = 0;

  function animate() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    // Dynamic Cloud Gradient Background
    const bgGrad = ctx.createLinearGradient(0, 0, 0, canvas.height);
    if (rainfall > 20) {
      bgGrad.addColorStop(0, '#CBD5E1');
      bgGrad.addColorStop(1, '#94A3B8');
    } else {
      bgGrad.addColorStop(0, '#E2E8F0');
      bgGrad.addColorStop(1, '#F1F5F9');
    }
    ctx.fillStyle = bgGrad;
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    // Lightning Trigger for Heavy Rain (>20mm)
    if (rainfall > 20 && Math.random() < 0.008) {
      flash = 1;
    }

    if (flash > 0) {
      ctx.fillStyle = `rgba(255, 255, 255, ${flash})`;
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      flash -= 0.05;
    }

    // Render Rain Particles
    ctx.strokeStyle = 'rgba(58, 111, 247, 0.45)';
    ctx.lineWidth = 1.2;
    ctx.beginPath();
    drops.forEach(d => {
      ctx.moveTo(d.x, d.y);
      ctx.lineTo(d.x - (rainfall > 10 ? 2 : 0), d.y + d.length);
      d.y += d.speed;
      d.x -= (rainfall > 10 ? 0.5 : 0);
      if (d.y > canvas.height) {
        d.y = -d.length;
        d.x = Math.random() * canvas.width;
      }
    });
    ctx.stroke();

    rainAnimationId = requestAnimationFrame(animate);
  }
  animate();
}

// 2. Wind Streamlines Visualization
let windAnimationId;
function initWindCanvas(windSpeed) {
  const canvas = document.getElementById('windCanvas');
  const ctx = canvas.getContext('2d');
  canvas.width = canvas.offsetWidth;
  canvas.height = canvas.offsetHeight;

  cancelAnimationFrame(windAnimationId);

  const particles = [];
  const count = 40;
  
  for (let i = 0; i < count; i++) {
    particles.push({
      x: Math.random() * canvas.width,
      y: Math.random() * canvas.height,
      length: Math.random() * 40 + 20,
      speed: (windSpeed / 5) + Math.random() * 1.5 + 0.8
    });
  }

  function animate() {
    ctx.fillStyle = '#F8FAFC';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    ctx.strokeStyle = 'rgba(100, 116, 139, 0.25)';
    ctx.lineWidth = 1.8;
    ctx.beginPath();

    particles.forEach(p => {
      ctx.moveTo(p.x, p.y);
      ctx.bezierCurveTo(
        p.x + p.length * 0.3, p.y - 4,
        p.x + p.length * 0.6, p.y + 4,
        p.x + p.length, p.y
      );
      p.x += p.speed;
      if (p.x > canvas.width + p.length) {
        p.x = -p.length;
        p.y = Math.random() * canvas.height;
      }
    });
    ctx.stroke();

    windAnimationId = requestAnimationFrame(animate);
  }
  animate();
}

// 3. Visibility Layered Horizon
function initVisCanvas(vis) {
  const canvas = document.getElementById('visCanvas');
  const ctx = canvas.getContext('2d');
  canvas.width = canvas.offsetWidth;
  canvas.height = canvas.offsetHeight;

  ctx.fillStyle = '#F1F5F9';
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  // Calculate Haze / Fog Opacity based on distance
  const fogOpacity = Math.max(0, 1 - (vis / 15));

  // Layered Mountain Range Vectors
  function drawDistantRidge(yOffset, color) {
    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.moveTo(0, canvas.height);
    ctx.lineTo(0, yOffset);
    ctx.quadraticCurveTo(canvas.width * 0.25, yOffset - 30, canvas.width * 0.5, yOffset);
    ctx.quadraticCurveTo(canvas.width * 0.75, yOffset + 20, canvas.width, yOffset - 10);
    ctx.lineTo(canvas.width, canvas.height);
    ctx.closePath();
    ctx.fill();
  }

  drawDistantRidge(130, '#CBD5E1');
  drawDistantRidge(170, '#94A3B8');

  // Atmospheric Fog Overlay
  if (fogOpacity > 0) {
    ctx.fillStyle = `rgba(248, 250, 252, ${fogOpacity * 0.85})`;
    ctx.fillRect(0, 0, canvas.width, canvas.height);
  }
}

/* ------------------------------------------------------------------
   APPLE STOCKS-INSPIRED SPARKLINE CHART
------------------------------------------------------------------ */

function initTrendChart(dates, values) {
  const canvas = document.getElementById('trendChart');
  const ctx = canvas.getContext('2d');
  const container = document.getElementById('chartContainer');
  
  canvas.width = container.offsetWidth;
  canvas.height = container.offsetHeight;

  const padding = 20;
  const w = canvas.width - padding * 2;
  const h = canvas.height - padding * 2;

  const maxVal = Math.max(...values.filter(v => v !== null), 10);
  const minVal = 0;

  const points = values.map((v, i) => {
    if (v === null) return null;
    const x = padding + (i / (values.length - 1)) * w;
    const y = canvas.height - padding - ((v - minVal) / (maxVal - minVal)) * h;
    return { x, y, val: v, date: dates[i] };
  }).filter(p => p !== null);

  if (points.length < 2) return;

  ctx.clearRect(0, 0, canvas.width, canvas.height);

  // Gradient Fill under curve
  const fillGrad = ctx.createLinearGradient(0, 0, 0, canvas.height);
  fillGrad.addColorStop(0, 'rgba(58, 111, 247, 0.2)');
  fillGrad.addColorStop(1, 'rgba(58, 111, 247, 0.0)');

  ctx.beginPath();
  ctx.moveTo(points[0].x, points[0].y);
  for (let i = 0; i < points.length - 1; i++) {
    const xc = (points[i].x + points[i + 1].x) / 2;
    const yc = (points[i].y + points[i + 1].y) / 2;
    ctx.quadraticCurveTo(points[i].x, points[i].y, xc, yc);
  }
  ctx.lineTo(points[points.length - 1].x, points[points.length - 1].y);
  ctx.lineTo(points[points.length - 1].x, canvas.height);
  ctx.lineTo(points[0].x, canvas.height);
  ctx.closePath();
  ctx.fillStyle = fillGrad;
  ctx.fill();

  // Draw Smooth Curve Line
  ctx.beginPath();
  ctx.moveTo(points[0].x, points[0].y);
  for (let i = 0; i < points.length - 1; i++) {
    const xc = (points[i].x + points[i + 1].x) / 2;
    const yc = (points[i].y + points[i + 1].y) / 2;
    ctx.quadraticCurveTo(points[i].x, points[i].y, xc, yc);
  }
  ctx.strokeStyle = '#3A6FF7';
  ctx.lineWidth = 2.5;
  ctx.stroke();
}

/* ------------------------------------------------------------------
   MAIN DASHBOARD RENDER
------------------------------------------------------------------ */

function renderDashboard() {
  if (!DATA) return;

  const station = stationSelect.value;
  const rainData = DATA.Rain[station];
  const windData = DATA.Wind[station];
  const visData = DATA.Vis[station];

  // Latest Available Telemetry Metrics
  const latestRain = rainData.grid.flat().filter(v => v !== null).pop() || 0;
  const latestWind = windData.grid.flat().filter(v => v !== null).pop() || 0;
  const latestVis = visData.grid.flat().filter(v => v !== null).pop() || 10;

  // Render Display Values
  document.getElementById('rainVal').innerHTML = `${latestRain} <span style="font-size:20px">mm</span>`;
  document.getElementById('windVal').innerHTML = `${latestWind} <span style="font-size:20px">km/h</span>`;
  document.getElementById('visVal').innerHTML = `${latestVis} <span style="font-size:20px">km</span>`;

  // Update Contextual Status Descriptors
  document.getElementById('rainStatus').textContent = 
    latestRain > 20 ? "Heavy Monsoon Activity" :
    latestRain > 5  ? "Moderate Precipitation" : "Light Precipitation";

  document.getElementById('windStatus').textContent = 
    latestWind > 25 ? "Gale Force Streams" :
    latestWind > 10 ? "Moderate Airflow" : "Gentle Flow";

  document.getElementById('visStatus').textContent = 
    latestVis < 5  ? "Dense Fog Haze" :
    latestVis < 10 ? "Moderate Haze" : "Optimal Conditions";

  // Trigger Dynamic Atmospheric Canvases
  initRainCanvas(latestRain);
  initWindCanvas(latestWind);
  initVisCanvas(latestVis);

  // Render Historical Trend Line
  const dailyRainTotals = rainData.grid.map(row => {
    const valid = row.filter(v => v !== null);
    return valid.length ? valid.reduce((a,b) => a+b, 0) : null;
  });
  initTrendChart(rainData.dates, dailyRainTotals);
}

stationSelect.addEventListener('change', renderDashboard);
window.addEventListener('resize', renderDashboard);
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
