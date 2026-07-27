"""
TMD Phuket Weather -> raw Excel log + immersive Apple-Weather-style HTML dashboard.

Excel (Phuket_Weather.xlsx) stores RAW DATA ONLY. All dynamic/interactive/visual experience
lives in Phuket_Weather_Dashboard.html, which fetches the xlsx file directly over HTTP(S)
when opened and renders an atmospheric, glassmorphic monitoring interface.
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

_HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Phuket Atmosphere</title>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
  :root{
    --primary:#3A6FF7;
    --primary-soft:rgba(58,111,247,0.12);
    --cloud-gray:#E8EDF3;
    --bg:#F8FAFC;
    --card-bg:rgba(255,255,255,0.62);
    --card-border:rgba(16,24,40,0.06);
    --text:#101418;
    --text-muted:rgba(16,20,24,0.52);
    --success:#1FA971;
    --warning:#F59E0B;
    --danger:#EF4444;
  }
  *{box-sizing:border-box;margin:0;padding:0;}
  html,body{
    background:var(--bg);
    color:var(--text);
    font-family:-apple-system,BlinkMacSystemFont,"SF Pro Display","SF Pro Text","Segoe UI",Roboto,sans-serif;
    -webkit-font-smoothing:antialiased;
    min-height:100vh;
  }
  .app{
    max-width:1180px;
    margin:0 auto;
    padding:36px 28px 80px;
  }

  /* ---------------- Top bar ---------------- */
  .topbar{
    display:flex;
    justify-content:space-between;
    align-items:flex-start;
    flex-wrap:wrap;
    gap:20px;
    margin-bottom:28px;
  }
  .eyebrow{
    font-size:12px;
    font-weight:600;
    letter-spacing:1.2px;
    text-transform:uppercase;
    color:var(--primary);
    margin-bottom:6px;
  }
  .location-block h1{
    font-size:34px;
    font-weight:650;
    letter-spacing:-0.5px;
    margin-bottom:6px;
  }
  .meta-row{
    display:flex;
    align-items:center;
    gap:10px;
    font-size:14px;
    color:var(--text-muted);
    font-weight:500;
  }
  .meta-row .dot{opacity:0.4;}

  .station-switch{
    display:flex;
    background:var(--cloud-gray);
    border-radius:999px;
    padding:4px;
    gap:2px;
  }
  .station-switch button{
    border:none;
    background:transparent;
    padding:9px 18px;
    border-radius:999px;
    font-size:13px;
    font-weight:600;
    color:var(--text-muted);
    cursor:pointer;
    transition:background 0.25s ease, color 0.25s ease, box-shadow 0.25s ease;
    font-family:inherit;
  }
  .station-switch button.active{
    background:#fff;
    color:var(--text);
    box-shadow:0 2px 10px rgba(16,24,40,0.10);
  }

  /* ---------------- Hero rain card ---------------- */
  .hero-card{
    position:relative;
    height:400px;
    border-radius:28px;
    overflow:hidden;
    box-shadow:0 24px 48px rgba(16,24,40,0.10);
    background:#dfe6ee;
  }
  #rainCanvas{
    position:absolute;
    inset:0;
    width:100%;
    height:100%;
    display:block;
  }
  .hero-card .lightning-flash{
    position:absolute;
    inset:0;
    background:#ffffff;
    opacity:0;
    pointer-events:none;
    transition:opacity 0.08s ease-out;
  }
  .glass-readout{
    position:absolute;
    left:26px;
    bottom:26px;
    padding:22px 26px;
    border-radius:22px;
    background:rgba(255,255,255,0.55);
    backdrop-filter:blur(24px) saturate(180%);
    -webkit-backdrop-filter:blur(24px) saturate(180%);
    border:1px solid rgba(255,255,255,0.5);
    box-shadow:0 8px 30px rgba(16,24,40,0.12);
    min-width:220px;
  }
  .glass-readout .label{
    font-size:13px;
    font-weight:600;
    color:var(--text-muted);
    text-transform:uppercase;
    letter-spacing:0.6px;
    margin-bottom:6px;
  }
  .glass-readout .value{
    font-size:40px;
    font-weight:650;
    letter-spacing:-1px;
    line-height:1;
  }
  .glass-readout .value .unit{
    font-size:16px;
    font-weight:600;
    color:var(--text-muted);
    margin-left:6px;
  }
  .glass-readout .descriptor{
    margin-top:8px;
    font-size:14px;
    font-weight:550;
    color:var(--text);
    opacity:0.85;
  }

  /* ---------------- Row cards ---------------- */
  .row-cards{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:22px;
    margin-top:22px;
  }
  .glass-card{
    position:relative;
    height:260px;
    border-radius:24px;
    overflow:hidden;
    box-shadow:0 18px 36px rgba(16,24,40,0.08);
  }
  .wind-card{background:#eef1f6;}
  #windCanvas{position:absolute;inset:0;width:100%;height:100%;display:block;}

  .vis-card{background:#eaf1f7;}
  .horizon-scene{position:absolute;inset:0;overflow:hidden;}
  .horizon-scene .sky{
    position:absolute;inset:0;
    background:linear-gradient(180deg,#cfe0ef 0%,#eef4f9 60%,#f8fbfd 100%);
  }
  .mountain{
    position:absolute;
    left:-10%;
    right:-10%;
    bottom:0;
    transition:opacity 0.6s ease, filter 0.6s ease, transform 0.6s ease;
  }
  .layer-3{
    height:55%;
    background:#aebdd0;
    clip-path:polygon(0% 100%, 0% 60%, 12% 40%, 26% 55%, 40% 30%, 55% 52%, 68% 25%, 82% 48%, 100% 35%, 100% 100%);
    opacity:0.55;
  }
  .layer-2{
    height:42%;
    background:#8fa2ba;
    clip-path:polygon(0% 100%, 0% 68%, 15% 45%, 30% 62%, 46% 38%, 60% 60%, 75% 34%, 90% 58%, 100% 46%, 100% 100%);
    opacity:0.72;
  }
  .layer-1{
    height:28%;
    background:#5f7592;
    clip-path:polygon(0% 100%, 0% 72%, 18% 50%, 34% 68%, 50% 46%, 66% 66%, 82% 42%, 100% 62%, 100% 100%);
    opacity:0.92;
  }
  .fog-wall{
    position:absolute;inset:0;
    background:#ffffff;
    opacity:0;
    backdrop-filter:blur(0px);
    -webkit-backdrop-filter:blur(0px);
    transition:opacity 0.6s ease, backdrop-filter 0.6s ease;
    pointer-events:none;
  }

  .card-readout{
    position:absolute;
    left:20px;
    bottom:20px;
    padding:16px 20px;
    border-radius:18px;
    background:rgba(255,255,255,0.55);
    backdrop-filter:blur(20px) saturate(180%);
    -webkit-backdrop-filter:blur(20px) saturate(180%);
    border:1px solid rgba(255,255,255,0.5);
    box-shadow:0 6px 22px rgba(16,24,40,0.10);
    min-width:190px;
  }
  .card-readout .label{
    font-size:12px;font-weight:600;color:var(--text-muted);
    text-transform:uppercase;letter-spacing:0.6px;margin-bottom:4px;
  }
  .card-readout .value{font-size:28px;font-weight:650;letter-spacing:-0.5px;line-height:1;}
  .card-readout .value .unit{font-size:13px;font-weight:600;color:var(--text-muted);margin-left:5px;}
  .card-readout .descriptor{margin-top:6px;font-size:13px;font-weight:550;color:var(--text);opacity:0.85;}

  /* ---------------- Trends ---------------- */
  .trends{margin-top:44px;}
  .trends h2{
    font-size:20px;font-weight:650;letter-spacing:-0.3px;margin-bottom:18px;
  }
  .trend-grid{
    display:grid;
    grid-template-columns:repeat(3, 1fr);
    gap:20px;
  }
  .trend-card{
    position:relative;
    background:var(--card-bg);
    backdrop-filter:blur(18px) saturate(160%);
    -webkit-backdrop-filter:blur(18px) saturate(160%);
    border:1px solid var(--card-border);
    border-radius:20px;
    padding:18px 20px 12px;
    box-shadow:0 10px 26px rgba(16,24,40,0.06);
  }
  .trend-card .t-head{
    display:flex;justify-content:space-between;align-items:baseline;margin-bottom:8px;
  }
  .trend-card .t-title{font-size:14px;font-weight:650;color:var(--text);}
  .trend-card .t-latest{font-size:13px;font-weight:600;color:var(--text-muted);}
  .trend-card canvas{width:100%;height:130px;display:block;}
  .trend-tooltip{
    position:absolute;
    top:14px;
    right:16px;
    font-size:12px;
    font-weight:600;
    padding:4px 10px;
    border-radius:8px;
    background:rgba(16,24,40,0.85);
    color:#fff;
    opacity:0;
    transform:translateY(-4px);
    transition:opacity 0.15s ease;
    pointer-events:none;
  }

  .file-status{
    position:fixed;
    right:22px;
    bottom:16px;
    font-size:12px;
    color:var(--text-muted);
    background:rgba(255,255,255,0.7);
    backdrop-filter:blur(10px);
    padding:6px 14px;
    border-radius:999px;
    border:1px solid var(--card-border);
  }
  .file-status.ok{color:var(--success);}
  .file-status.error{color:var(--danger);}

  @media (max-width:820px){
    .row-cards{grid-template-columns:1fr;}
    .trend-grid{grid-template-columns:1fr;}
    .hero-card{height:340px;}
  }
</style>
</head>
<body>
  <div class="app">
    <header class="topbar">
      <div class="location-block">
        <div class="eyebrow">Live Atmospheric Monitoring</div>
        <h1 id="locationName">Phuket</h1>
        <div class="meta-row">
          <span id="timestamp">—</span>
          <span class="dot">•</span>
          <span id="condition">Connecting to live data…</span>
        </div>
      </div>
      <div class="station-switch" id="stationSwitch"></div>
    </header>

    <section class="hero-card" id="rainHero">
      <canvas id="rainCanvas"></canvas>
      <div class="lightning-flash" id="lightningFlash"></div>
      <div class="glass-readout">
        <div class="label">Rainfall</div>
        <div class="value"><span id="rainValue">--</span><span class="unit">mm</span></div>
        <div class="descriptor" id="rainDescriptor">Awaiting data</div>
      </div>
    </section>

    <section class="row-cards">
      <div class="glass-card wind-card" id="windCard">
        <canvas id="windCanvas"></canvas>
        <div class="card-readout">
          <div class="label">Wind Speed</div>
          <div class="value"><span id="windValue">--</span><span class="unit">km/h</span></div>
          <div class="descriptor" id="windDescriptor">—</div>
        </div>
      </div>

      <div class="glass-card vis-card" id="visCard">
        <div class="horizon-scene">
          <div class="sky"></div>
          <div class="mountain layer-3" id="layer3"></div>
          <div class="mountain layer-2" id="layer2"></div>
          <div class="mountain layer-1" id="layer1"></div>
          <div class="fog-wall" id="fogWall"></div>
        </div>
        <div class="card-readout">
          <div class="label">Visibility</div>
          <div class="value"><span id="visValue">--</span><span class="unit">km</span></div>
          <div class="descriptor" id="visDescriptor">—</div>
        </div>
      </div>
    </section>

    <section class="trends">
      <h2>Historical Trends</h2>
      <div class="trend-grid" id="trendGrid">
        <div class="trend-card">
          <div class="trend-tooltip" id="tipRain"></div>
          <div class="t-head"><span class="t-title">Rainfall</span><span class="t-latest" id="latestRain">—</span></div>
          <canvas id="chartRain"></canvas>
        </div>
        <div class="trend-card">
          <div class="trend-tooltip" id="tipWind"></div>
          <div class="t-head"><span class="t-title">Wind Speed</span><span class="t-latest" id="latestWind">—</span></div>
          <canvas id="chartWind"></canvas>
        </div>
        <div class="trend-card">
          <div class="trend-tooltip" id="tipVis"></div>
          <div class="t-head"><span class="t-title">Visibility</span><span class="t-latest" id="latestVis">—</span></div>
          <canvas id="chartVis"></canvas>
        </div>
      </div>
    </section>
  </div>

  <div id="fileStatus" class="file-status">Initializing…</div>

<script>
const CONFIG = __CONFIG_JSON__;

/* ============================================================
   Classification tables
============================================================ */
const RAIN_LEVELS = [
  { max:0,        key:'dry',      label:'Dry Conditions',        sub:'No rainfall detected',            particles:0,   speed:[0,0],   sky:['#cfd8e3','#eef2f6'] },
  { max:5,        key:'light',    label:'Light Rain',            sub:'Sparse showers passing through',  particles:70,  speed:[2.2,3.6], sky:['#b7c5d7','#dee6ee'] },
  { max:20,       key:'moderate', label:'Moderate Rain',         sub:'Steady rainfall over the area',   particles:170, speed:[4.5,7],   sky:['#8b9db6','#c1cad7'] },
  { max:50,       key:'heavy',    label:'Heavy Rain',            sub:'Intense monsoon activity',        particles:320, speed:[8,12.5],  sky:['#54647c','#8994a6'] },
  { max:Infinity, key:'extreme',  label:'Severe Thunderstorm',   sub:'Extreme rainfall — lightning risk', particles:460, speed:[13,19],  sky:['#1e2531','#454f60'] },
];
function classifyRain(mm){
  const v = (mm===null||mm===undefined||isNaN(mm)) ? 0 : mm;
  for (const l of RAIN_LEVELS) if (v<=l.max) return l;
  return RAIN_LEVELS[RAIN_LEVELS.length-1];
}

const WIND_LEVELS = [
  { max:10,       key:'calm',    label:'Calm Air',        particles:40,  speed:[0.4,0.8], amp:4 },
  { max:25,       key:'breezy',  label:'Gentle Breeze',   particles:90,  speed:[1.0,1.8], amp:8 },
  { max:40,       key:'windy',   label:'Strong Flow',     particles:150, speed:[2.0,3.2], amp:14 },
  { max:Infinity, key:'gale',    label:'Turbulent Gale',  particles:220, speed:[3.4,4.8], amp:22 },
];
function classifyWind(v){
  const s = (v===null||v===undefined||isNaN(v)) ? 0 : v;
  for (const l of WIND_LEVELS) if (s<=l.max) return l;
  return WIND_LEVELS[WIND_LEVELS.length-1];
}

const VIS_LEVELS = [
  { min:10,   key:'excellent', label:'Clear Conditions',   fog:0.02, blur:0 },
  { min:5,    key:'moderate',  label:'Slight Haze',        fog:0.28, blur:2 },
  { min:2,    key:'poor',      label:'Reduced Visibility', fog:0.58, blur:6 },
  { min:-1,   key:'verypoor',  label:'Dense Fog',          fog:0.86, blur:12 },
];
function classifyVisibility(v){
  const s = (v===null||v===undefined||isNaN(v)) ? 0 : v;
  for (const l of VIS_LEVELS) if (s>=l.min) return l;
  return VIS_LEVELS[VIS_LEVELS.length-1];
}

/* ============================================================
   State
============================================================ */
let recordsByStation = {};
let currentStation = 'COMBINED';

/* ============================================================
   Station switch UI
============================================================ */
const switchEl = document.getElementById('stationSwitch');
CONFIG.stations.forEach(s=>{
  const btn = document.createElement('button');
  btn.textContent = s.label_display;
  btn.dataset.code = s.code;
  if (s.code === currentStation) btn.classList.add('active');
  btn.addEventListener('click', ()=>{
    currentStation = s.code;
    document.querySelectorAll('.station-switch button').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    renderAll();
  });
  switchEl.appendChild(btn);
});

function stationCfg(code){ return CONFIG.stations.find(s=>s.code===code); }

/* ============================================================
   Data loading
============================================================ */
function setStatus(msg, cls){
  const el = document.getElementById('fileStatus');
  el.textContent = msg;
  el.className = 'file-status ' + (cls||'');
}

function parseDT(v){
  if (v===null||v===undefined||v==='') return null;
  const d = (v instanceof Date) ? v : new Date(v);
  return isNaN(d.getTime()) ? null : d;
}
function numOrNull(v){
  if (v===null||v===undefined||v==='') return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}

async function loadData(){
  setStatus('Syncing live data…');
  try{
    const res = await fetch(CONFIG.xlsx_filename, {cache:'no-cache'});
    if (!res.ok) throw new Error('HTTP '+res.status);
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, {type:'array', cellDates:true});
    const sheet = wb.Sheets[CONFIG.data_sheet_name];
    if (!sheet) throw new Error('Sheet missing');
    const rows = XLSX.utils.sheet_to_json(sheet, {defval:null});

    const byStation = {};
    CONFIG.stations.forEach(s=> byStation[s.code] = []);

    rows.forEach(row=>{
      const dt = parseDT(row.DateTime);
      if (!dt || !row.Station) return;
      const cfg = CONFIG.stations.find(s=>s.label_raw === row.Station);
      if (!cfg) return;
      byStation[cfg.code].push({
        dt,
        Rain: numOrNull(row.Rainfall_mm),
        Wind: numOrNull(row.WindSpeed),
        Vis: numOrNull(row.LandVisibility),
      });
    });

    Object.keys(byStation).forEach(code=>{
      byStation[code].sort((a,b)=>a.dt-b.dt);
    });

    recordsByStation = byStation;
    setStatus('Live • updated ' + new Date().toLocaleTimeString(), 'ok');
    renderAll();
  }catch(err){
    setStatus('Unable to reach live dataset', 'error');
  }
}

/* ============================================================
   Formatting helpers
============================================================ */
function fmtDT(d){
  if (!d) return '—';
  const opts = {weekday:'short', month:'short', day:'numeric', hour:'2-digit', minute:'2-digit'};
  return d.toLocaleString('en-GB', opts);
}
function fmt1(v){ return (v===null||v===undefined||isNaN(v)) ? '--' : v.toFixed(1); }

function deriveCondition(rec){
  if (!rec) return 'No data available';
  const rl = classifyRain(rec.Rain);
  if (rl.key !== 'dry') return rl.label + ' · ' + classifyWind(rec.Wind).label;
  const vl = classifyVisibility(rec.Vis);
  if (vl.key !== 'excellent') return vl.label;
  return 'Clear & Calm';
}

/* ============================================================
   RAIN HERO ANIMATION
============================================================ */
const rainCanvas = document.getElementById('rainCanvas');
const rainCtx = rainCanvas.getContext('2d');
const lightningFlash = document.getElementById('lightningFlash');

let rainParticles = [];
const RAIN_MAX_PARTICLES = 500;
for (let i=0;i<RAIN_MAX_PARTICLES;i++){
  rainParticles.push({ x:Math.random(), y:Math.random(), len:10+Math.random()*18, phase:Math.random()*10 });
}

let rainState = { level: RAIN_LEVELS[0], visualCount:0, flashAlpha:0, boltLife:0, boltPts:[] };

function resizeCanvas(canvas){
  const rect = canvas.parentElement.getBoundingClientRect();
  const dpr = window.devicePixelRatio || 1;
  canvas.width = rect.width * dpr;
  canvas.height = rect.height * dpr;
  canvas.getContext('2d').setTransform(dpr,0,0,dpr,0,0);
}

function setRainScene(mm){
  const lvl = classifyRain(mm);
  rainState.level = lvl;
  document.getElementById('rainValue').textContent = (mm===null||mm===undefined) ? '--' : mm.toFixed(1);
  document.getElementById('rainDescriptor').textContent = lvl.label + ' — ' + lvl.sub;
}

function drawRainFrame(){
  const w = rainCanvas.clientWidth, h = rainCanvas.clientHeight;
  const lvl = rainState.level;

  // smooth particle count transition
  rainState.visualCount += (lvl.particles - rainState.visualCount) * 0.04;
  const count = Math.round(rainState.visualCount);

  // sky gradient
  const grad = rainCtx.createLinearGradient(0,0,0,h);
  grad.addColorStop(0, lvl.sky[0]);
  grad.addColorStop(1, lvl.sky[1]);
  rainCtx.fillStyle = grad;
  rainCtx.fillRect(0,0,w,h);

  // soft cloud shading blobs (slow drifting)
  const t = performance.now()*0.00006;
  rainCtx.globalAlpha = 0.14;
  for (let i=0;i<4;i++){
    const cx = ((Math.sin(t+i*1.7)+1)/2) * w;
    const cy = h*0.18 + i*22;
    const r = 140 + i*30;
    const cg = rainCtx.createRadialGradient(cx,cy,0,cx,cy,r);
    cg.addColorStop(0, 'rgba(255,255,255,0.55)');
    cg.addColorStop(1, 'rgba(255,255,255,0)');
    rainCtx.fillStyle = cg;
    rainCtx.beginPath();
    rainCtx.arc(cx,cy,r,0,Math.PI*2);
    rainCtx.fill();
  }
  rainCtx.globalAlpha = 1;

  // rain streaks
  if (count > 0){
    const [minS,maxS] = lvl.speed;
    rainCtx.strokeStyle = 'rgba(255,255,255,0.42)';
    rainCtx.lineWidth = 1.1;
    rainCtx.lineCap = 'round';
    for (let i=0;i<count;i++){
      const p = rainParticles[i];
      const speed = minS + (maxS-minS) * ((Math.sin(p.phase)+1)/2);
      p.y += (speed/h) * 0.9;
      if (p.y > 1.05){ p.y = -0.05; p.x = Math.random(); }
      const px = p.x*w, py = p.y*h;
      rainCtx.beginPath();
      rainCtx.moveTo(px, py);
      rainCtx.lineTo(px-3, py-p.len);
      rainCtx.stroke();
    }
  }

  // mist accumulation for heavy/extreme
  if (lvl.key==='heavy' || lvl.key==='extreme'){
    const mg = rainCtx.createLinearGradient(0,h*0.75,0,h);
    mg.addColorStop(0,'rgba(255,255,255,0)');
    mg.addColorStop(1,'rgba(255,255,255,0.22)');
    rainCtx.fillStyle = mg;
    rainCtx.fillRect(0,h*0.7,w,h*0.3);
  }

  // lightning logic
  if (lvl.key === 'extreme'){
    if (rainState.boltLife<=0 && Math.random() < 0.008){
      rainState.flashAlpha = 0.75;
      rainState.boltLife = 6;
      const pts = [];
