"""
GPS vs MIS Fleet Dashboard - Streamlit version
================================================
Same dashboard as the web app, but as a Streamlit app you run on your own
machine or company server. Gives you a URL that works on your company
network (Wi-Fi / LAN) without ever touching the public internet.

RUN:
    pip install -r requirements.txt
    streamlit run app.py

Streamlit will print two URLs:
    Local URL:   http://localhost:8501       (only this PC)
    Network URL: http://192.168.x.x:8501     (anyone on your company network)

Share the Network URL with your team - it only works for people on the same
office Wi-Fi / LAN, never reaches the public internet.
"""

import base64
import io
import json
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import load_workbook

st.set_page_config(page_title="GPS vs MIS Fleet Dashboard", layout="wide", page_icon="🚚", initial_sidebar_state="expanded")

# ---------------------------------------------------------------------------
# Theme — light and dark palettes. The active one follows Streamlit's own
# Settings -> Theme toggle (top-right menu / "Use system setting"), read via
# st.context.theme.type, so native widgets (dataframe grid, dropdowns) and
# our custom CSS + charts always agree on which mode is active.
#
# GPS / MIS series colours are a colour-blind-safe blue/orange pair; status
# colours (ok / watch / warn / critical) are kept separate from the series
# colours and always appear with a text label, never colour alone.
# ---------------------------------------------------------------------------
LIGHT_THEME = {
    "COLOR_BG": "#F4F6FA", "COLOR_CARD": "#FFFFFF", "COLOR_PANEL": "#FFFFFF", "COLOR_BORDER": "#E3E8EF",
    "COLOR_TEXT": "#0F172A", "COLOR_MUTED": "#64748B",
    "COLOR_GPS": "#2A78D6", "COLOR_MIS": "#EB6834", "COLOR_OK": "#15803D",
    "COLOR_CRIT": "#D03B3B", "COLOR_WARN": "#B45309", "COLOR_DIFF": "#6D28D9",
    "COLOR_VEHICLES": "#1E40AF", "COLOR_FLAGGED": "#D03B3B",
    "CHART_TEXT": "#475569", "CHART_GRID": "#EEF1F5", "CHART_AXIS": "#CBD5E1",
}
DARK_THEME = {
    "COLOR_BG": "#0B1120", "COLOR_CARD": "#121A2B", "COLOR_PANEL": "#0F1626", "COLOR_BORDER": "#223049",
    "COLOR_TEXT": "#F1F5F9", "COLOR_MUTED": "#94A3B8",
    "COLOR_GPS": "#3987E5", "COLOR_MIS": "#D95926", "COLOR_OK": "#22C55E",
    "COLOR_CRIT": "#E66767", "COLOR_WARN": "#FAB219", "COLOR_DIFF": "#A78BFA",
    "COLOR_VEHICLES": "#6E8BFF", "COLOR_FLAGGED": "#E66767",
    "CHART_TEXT": "#A8B3C7", "CHART_GRID": "#1C2740", "CHART_AXIS": "#334155",
}

_theme_type = getattr(st.context.theme, "type", None) or "light"
_palette = DARK_THEME if _theme_type == "dark" else LIGHT_THEME
_IS_DARK = _theme_type == "dark"

COLOR_BG = _palette["COLOR_BG"]
COLOR_CARD = _palette["COLOR_CARD"]
COLOR_PANEL = _palette["COLOR_PANEL"]
COLOR_BORDER = _palette["COLOR_BORDER"]
COLOR_TEXT = _palette["COLOR_TEXT"]
COLOR_MUTED = _palette["COLOR_MUTED"]
COLOR_GPS = _palette["COLOR_GPS"]
COLOR_MIS = _palette["COLOR_MIS"]
COLOR_OK = _palette["COLOR_OK"]
COLOR_CRIT = _palette["COLOR_CRIT"]
COLOR_WARN = _palette["COLOR_WARN"]
COLOR_DIFF = _palette["COLOR_DIFF"]
COLOR_VEHICLES = _palette["COLOR_VEHICLES"]
COLOR_FLAGGED = _palette["COLOR_FLAGGED"]
CHART_TEXT = _palette["CHART_TEXT"]
CHART_GRID = _palette["CHART_GRID"]
CHART_AXIS = _palette["CHART_AXIS"]
CHART_BG = COLOR_CARD

# Status fills used for chart marks (bars, donut slices). Fixed across themes.
STATUS_FILL = {"critical": "#D03B3B", "warn": "#EC835A", "watch": "#F2A818", "ok": "#0CA30C"}
STATUS_LABEL = {"critical": "Critical", "warn": "Needs check", "watch": "Watch", "ok": "OK"}

# Brand band behind the page header — deep corporate navy in both themes.
BRAND_GRADIENT = "linear-gradient(120deg, #0B1F4B 0%, #13317A 55%, #1D4ED8 100%)"

# Neutral (non-colored) ambient shadow — a plain black shadow reads fine on
# both a white and a near-black surface; only the alpha needs to differ.
SHADOW_SM = "rgba(0,0,0,0.35)" if _IS_DARK else "rgba(15,23,42,0.05)"
SHADOW_LG = "rgba(0,0,0,0.45)" if _IS_DARK else "rgba(15,23,42,0.07)"


def hex_to_rgba(hex_color, alpha):
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {{ font-family: 'Inter', system-ui, 'Segoe UI', sans-serif; font-size: 15px; line-height: 1.55; }}
.stApp {{ background: {COLOR_BG}; color: {COLOR_TEXT}; }}
footer {{ visibility: hidden; }}
.block-container {{ padding-top: 2.2rem !important; max-width: 1480px; }}

h1, h2, h3, h4 {{ font-family: 'Inter', system-ui, sans-serif !important; color: {COLOR_TEXT} !important; }}
h1 {{ font-weight: 700 !important; font-size: 30px !important; letter-spacing: -0.5px; }}
h2, h3 {{ font-weight: 650 !important; font-size: 19px !important; letter-spacing: -0.2px; margin-top: 10px !important; }}
h4 {{ font-weight: 600 !important; font-size: 16px !important; }}
p, label, span, div {{ font-size: 15px; font-weight: 400; }}

.eyebrow {{
    color: {COLOR_MUTED}; font-weight: 700; font-size: 12px;
    letter-spacing: 1.6px; text-transform: uppercase; margin-bottom: 4px;
}}

/* ---------- Brand header band ---------- */
.brand-hero {{
    background: {BRAND_GRADIENT}; border-radius: 18px; padding: 26px 30px;
    display: flex; justify-content: space-between; align-items: flex-end; gap: 20px; flex-wrap: wrap;
    position: relative; overflow: hidden; margin-bottom: 22px;
    box-shadow: 0 10px 30px rgba(11,31,75,0.25);
}}
.brand-hero::after {{
    content: ""; position: absolute; right: -60px; top: -80px; width: 280px; height: 280px;
    border-radius: 50%; background: radial-gradient(circle, rgba(255,255,255,0.14), rgba(255,255,255,0) 70%);
}}
.brand-hero .hero-eyebrow {{ color: rgba(255,255,255,0.72) !important; font-size: 12px !important; font-weight: 700 !important; letter-spacing: 1.6px; text-transform: uppercase; }}
.brand-hero .hero-title {{ color: #FFFFFF !important; font-size: 28px !important; font-weight: 750 !important; letter-spacing: -0.5px; line-height: 1.2; margin-top: 4px; }}
.brand-hero .hero-sub {{ color: rgba(255,255,255,0.78) !important; font-size: 14px !important; margin-top: 6px; }}
.brand-hero .hero-chips {{ display: flex; gap: 8px; flex-wrap: wrap; position: relative; z-index: 1; }}
.brand-hero .hero-chip {{
    display: inline-flex; align-items: center; gap: 6px; padding: 7px 12px; border-radius: 999px;
    background: rgba(255,255,255,0.12); border: 1px solid rgba(255,255,255,0.22);
    color: #FFFFFF !important; font-size: 13px !important; font-weight: 600 !important; white-space: nowrap;
}}
.brand-hero .hero-chip svg {{ width: 15px; height: 15px; }}

/* ---------- KPI cards ---------- */
.kpi-card {{
    background: {COLOR_CARD};
    border: 1px solid {COLOR_BORDER}; border-radius: 16px;
    padding: 18px 20px 16px; position: relative; overflow: hidden;
    box-shadow: 0 1px 2px {SHADOW_SM}, 0 8px 24px {SHADOW_LG};
    transition: transform 0.15s ease, box-shadow 0.15s ease, border-color 0.15s ease;
    display: flex; flex-direction: column; min-height: 220px; box-sizing: border-box;
}}
.kpi-card::before {{
    content: ""; position: absolute; left: 0; right: 0; top: 0; height: 3px; background: var(--accent);
}}
.kpi-card:hover {{
    transform: translateY(-2px);
    border-color: color-mix(in srgb, var(--accent) 40%, {COLOR_BORDER});
    box-shadow: 0 2px 4px {SHADOW_SM}, 0 14px 30px {SHADOW_LG};
}}
.kpi-card .kpi-head {{ display: flex; justify-content: space-between; align-items: center; gap: 8px; margin-bottom: 12px; }}
.kpi-card .kpi-label {{ color: {COLOR_MUTED} !important; font-size: 12px !important; letter-spacing: 0.6px; text-transform: uppercase; font-weight: 650 !important; }}
.kpi-card .kpi-icon {{
    width: 34px; height: 34px; border-radius: 10px; flex: none;
    display: flex; align-items: center; justify-content: center; color: var(--accent);
    background: color-mix(in srgb, var(--accent) 12%, transparent);
}}
.kpi-card .kpi-icon svg {{ width: 18px; height: 18px; }}
.kpi-card .kpi-value {{
    color: {COLOR_TEXT} !important; font-size: clamp(22px, 1.9vw, 30px) !important; font-weight: 750 !important;
    letter-spacing: -0.6px; line-height: 1.1; white-space: nowrap; margin-bottom: 10px;
}}
.kpi-card .kpi-unit {{ color: {COLOR_MUTED} !important; font-size: 13px !important; font-weight: 600 !important; margin-left: 5px; letter-spacing: 0; }}
.kpi-card .kpi-foot {{ display: flex; flex-direction: column; align-items: flex-start; gap: 6px; }}
.kpi-card .kpi-delta {{
    display: inline-flex; align-items: center; gap: 3px; font-size: 12px !important; font-weight: 650 !important;
    padding: 3px 8px; border-radius: 999px; white-space: nowrap;
}}
.kpi-card .kpi-vs {{ color: {COLOR_MUTED} !important; font-size: 12px !important; font-weight: 500 !important; white-space: nowrap; }}
.kpi-card .kpi-spark {{ margin-top: auto; padding-top: 12px; }}
.kpi-card .kpi-spark svg {{ width: 100%; height: 34px; display: block; }}
.kpi-card .kpi-meter {{ margin-top: auto; padding-top: 14px; }}
.kpi-card .kpi-meter-track {{ height: 6px; border-radius: 999px; background: color-mix(in srgb, var(--accent) 14%, transparent); overflow: hidden; }}
.kpi-card .kpi-meter-fill {{ height: 100%; border-radius: 999px; background: var(--accent); }}

/* ---------- Chart cards ---------- */
.card-title {{ color: {COLOR_TEXT} !important; font-size: 16px !important; font-weight: 650 !important; letter-spacing: -0.2px; }}
.card-sub {{ color: {COLOR_MUTED} !important; font-size: 13px !important; margin-top: 2px; margin-bottom: 4px; }}
.legend-row {{ display: flex; gap: 16px; flex-wrap: wrap; margin-top: 6px; }}
.legend-item {{ display: inline-flex; align-items: center; gap: 6px; color: {COLOR_MUTED} !important; font-size: 12.5px !important; font-weight: 500 !important; }}
.legend-swatch {{ width: 10px; height: 10px; border-radius: 3px; display: inline-block; }}

.action-card {{
    background: {COLOR_CARD}; border: 1px solid {COLOR_BORDER}; border-radius: 12px;
    padding: 14px 18px; margin-bottom: 10px; box-shadow: 0 1px 2px {SHADOW_SM};
}}

/* Chart cards — bordered st.container(key="card_...") boxes */
[class*="st-key-card_"] {{
    background: {COLOR_CARD} !important; border: 1px solid {COLOR_BORDER} !important;
    border-radius: 16px !important; padding: 20px 20px 12px !important;
    box-shadow: 0 1px 2px {SHADOW_SM}, 0 8px 24px {SHADOW_LG};
}}

/* Dataframe / tables */
.stDataFrame, [data-testid="stDataFrame"] {{
    border: 1px solid {COLOR_BORDER} !important; border-radius: 12px !important; overflow: hidden;
    box-shadow: 0 1px 2px {SHADOW_SM};
}}

/* Inputs */
.stTextInput input, .stSelectbox [data-baseweb="select"] > div, .stTextInput > div > div {{
    background: {COLOR_CARD} !important; border: 1px solid {COLOR_BORDER} !important;
    color: {COLOR_TEXT} !important; border-radius: 8px !important; font-size: 14px !important;
}}
.stSlider [data-baseweb="slider"] {{ padding-top: 6px; }}
.stSlider [role="slider"] {{ background: {COLOR_VEHICLES} !important; box-shadow: 0 0 0 5px {hex_to_rgba(COLOR_VEHICLES, 0.15)} !important; }}
.stSlider div[data-baseweb="slider"] > div > div {{ background: {COLOR_VEHICLES} !important; }}

/* Widget labels (Viewing, Flag threshold, Search, etc.) */
[data-testid="stWidgetLabel"] p, [data-testid="stWidgetLabel"] label {{
    font-weight: 600 !important; color: {COLOR_TEXT} !important; font-size: 14px !important;
}}
[data-testid="stMarkdownContainer"] p {{ font-weight: 400; color: {COLOR_TEXT}; }}

/* File uploader */
[data-testid="stFileUploaderDropzone"] {{
    background: {COLOR_BG} !important;
    border: 1.5px dashed {COLOR_BORDER} !important; border-radius: 12px !important;
    transition: border-color 0.15s ease, box-shadow 0.15s ease;
}}
[data-testid="stFileUploaderDropzone"]:hover {{ border-color: {COLOR_VEHICLES} !important; box-shadow: 0 0 0 3px {hex_to_rgba(COLOR_VEHICLES, 0.10)} !important; }}

/* Buttons — normal size by default */
.stButton button, .stFormSubmitButton button {{
    background: {COLOR_VEHICLES} !important; color: #FFFFFF !important; border: none !important;
    border-radius: 8px !important; font-weight: 600 !important; font-size: 14px !important;
    padding: 0.4rem 1rem !important; transition: background 0.15s ease, transform 0.1s ease, box-shadow 0.15s ease;
    box-shadow: 0 1px 2px {SHADOW_SM};
}}
.stButton button:hover, .stFormSubmitButton button:hover {{
    background: color-mix(in srgb, {COLOR_VEHICLES} 88%, black) !important;
    transform: translateY(-1px); box-shadow: 0 4px 10px {SHADOW_LG};
}}

.stButton button p, .stFormSubmitButton button p {{ color: #FFFFFF !important; font-weight: 600 !important; font-size: 14px !important; }}

/* Corrective-action tiles — soft status cards, not solid buttons */
.st-key-corrective_actions button {{
    white-space: pre-line !important; line-height: 1.35 !important; min-height: 64px !important;
    font-size: 13px !important; font-weight: 600 !important;
    border-radius: 12px !important; text-align: left !important;
    box-shadow: 0 1px 2px {SHADOW_SM} !important;
}}

/* Corrective-action tiles colored by severity */
[class*="st-key-actsev_critical"] button {{
    background: color-mix(in srgb, {COLOR_CRIT} 7%, {COLOR_CARD}) !important;
    border: 1px solid color-mix(in srgb, {COLOR_CRIT} 28%, {COLOR_BORDER}) !important;
    border-left: 4px solid {COLOR_CRIT} !important; color: {COLOR_CRIT} !important;
}}
[class*="st-key-actsev_warn"] button {{
    background: color-mix(in srgb, {COLOR_WARN} 7%, {COLOR_CARD}) !important;
    border: 1px solid color-mix(in srgb, {COLOR_WARN} 28%, {COLOR_BORDER}) !important;
    border-left: 4px solid {COLOR_WARN} !important; color: {COLOR_WARN} !important;
}}
[class*="st-key-actsev_watch"] button {{
    background: color-mix(in srgb, {COLOR_WARN} 4%, {COLOR_CARD}) !important;
    border: 1px solid color-mix(in srgb, {COLOR_WARN} 18%, {COLOR_BORDER}) !important;
    border-left: 4px solid color-mix(in srgb, {COLOR_WARN} 55%, {COLOR_BORDER}) !important; color: {COLOR_WARN} !important;
}}
[class*="st-key-actsev_ok"] button {{
    background: color-mix(in srgb, {COLOR_OK} 6%, {COLOR_CARD}) !important;
    border: 1px solid color-mix(in srgb, {COLOR_OK} 25%, {COLOR_BORDER}) !important;
    border-left: 4px solid {COLOR_OK} !important; color: {COLOR_OK} !important;
}}
.st-key-corrective_actions button:hover {{ transform: translateY(-1px); }}
.st-key-corrective_actions button p {{ color: inherit !important; font-size: 13px !important; text-align: left !important; }}

/* Key findings panel */
.insight-card {{
    background: {COLOR_CARD}; border: 1px solid {COLOR_BORDER}; border-radius: 16px;
    box-shadow: 0 1px 2px {SHADOW_SM}, 0 8px 24px {SHADOW_LG};
    display: flex; gap: 28px; padding: 22px 26px; margin: 6px 0 18px; flex-wrap: wrap;
}}
.insight-card .score {{ display: flex; flex-direction: column; align-items: center; justify-content: center; min-width: 170px; text-align: center; }}
.insight-card .score-ring {{
    width: 128px; height: 128px; border-radius: 50%;
    background: conic-gradient(var(--c) calc(var(--p) * 1%), color-mix(in srgb, var(--c) 14%, transparent) 0);
    display: flex; align-items: center; justify-content: center;
}}
.insight-card .score-inner {{ width: 100px; height: 100px; border-radius: 50%; background: {COLOR_CARD}; display: flex; flex-direction: column; align-items: center; justify-content: center; }}
.insight-card .score-val {{ color: {COLOR_TEXT} !important; font-size: 28px !important; font-weight: 750 !important; line-height: 1.1; }}
.insight-card .score-lbl {{ color: {COLOR_MUTED} !important; font-size: 12px !important; font-weight: 500 !important; }}
.insight-card .score-cap {{ color: {COLOR_TEXT} !important; font-size: 14px !important; font-weight: 650 !important; margin-top: 12px; }}
.insight-card .score-sub {{ color: {COLOR_MUTED} !important; font-size: 12.5px !important; max-width: 180px; }}
.insight-card .insight-body {{ flex: 1; min-width: 280px; border-left: 1px solid {COLOR_BORDER}; padding-left: 28px; }}
.insight-card .insight-title {{ color: {COLOR_TEXT} !important; font-size: 16px !important; font-weight: 650 !important; margin-bottom: 10px; }}
.insight-card .insight-row {{ display: flex; gap: 10px; align-items: flex-start; padding: 7px 0; border-bottom: 1px dashed {COLOR_BORDER}; }}
.insight-card .insight-row:last-child {{ border-bottom: none; }}
.insight-card .insight-dot {{ width: 9px; height: 9px; border-radius: 50%; flex: none; margin-top: 7px; }}
.insight-card .insight-text {{ color: {COLOR_TEXT} !important; font-size: 14px !important; line-height: 1.5; }}
.insight-card .insight-text b {{ font-weight: 650; }}

/* Download button */
.stDownloadButton button {{
    background: {COLOR_CARD} !important; color: {COLOR_VEHICLES} !important;
    border: 1px solid color-mix(in srgb, {COLOR_VEHICLES} 45%, {COLOR_BORDER}) !important;
    border-radius: 8px !important; font-weight: 600 !important; box-shadow: 0 1px 2px {SHADOW_SM};
}}
.stDownloadButton button p {{ color: {COLOR_VEHICLES} !important; font-weight: 600 !important; font-size: 14px !important; }}
.stDownloadButton button:hover {{ background: color-mix(in srgb, {COLOR_VEHICLES} 8%, {COLOR_CARD}) !important; }}

/* Footer */
.app-footer {{
    margin-top: 36px; padding: 18px 4px 6px; border-top: 1px solid {COLOR_BORDER};
    display: flex; justify-content: space-between; gap: 12px; flex-wrap: wrap;
}}
.app-footer span {{ color: {COLOR_MUTED} !important; font-size: 12.5px !important; }}
.app-footer b {{ color: {COLOR_TEXT}; font-weight: 600; }}

/* Sidebar */
[data-testid="stSidebar"] {{ background: {COLOR_PANEL} !important; border-right: 1px solid {COLOR_BORDER}; }}
[data-testid="stSidebar"] img {{ border-radius: 8px; }}

/* Captions */
.stCaption, [data-testid="stCaptionContainer"] {{ color: {COLOR_MUTED} !important; font-weight: 500 !important; font-size: 13px !important; }}

/* Checkbox label */
.stCheckbox label p {{ color: {COLOR_TEXT} !important; font-size: 14px !important; }}

/* Alert boxes */
[data-testid="stAlert"] {{ background: {COLOR_CARD} !important; border: 1px solid {COLOR_BORDER} !important; border-radius: 10px; }}
</style>
""", unsafe_allow_html=True)

ACTION_COLORS = {"critical": COLOR_CRIT, "warn": COLOR_WARN, "watch": COLOR_WARN, "ok": COLOR_OK}


# Inline line icons (stroke = currentColor) — crisper and more corporate than emoji.
_ICON_PATHS = {
    "gps": '<polygon points="3 11 22 2 13 21 11 13 3 11"/>',
    "mis": '<rect x="8" y="2" width="8" height="4" rx="1"/><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><path d="M9 12h6"/><path d="M9 16h4"/>',
    "diff": '<path d="M8 3 4 7l4 4"/><path d="M4 7h16"/><path d="m16 21 4-4-4-4"/><path d="M20 17H4"/>',
    "truck": '<path d="M14 18V6a2 2 0 0 0-2-2H4a2 2 0 0 0-2 2v11a1 1 0 0 0 1 1h2"/><path d="M15 18H9"/><path d="M19 18h2a1 1 0 0 0 1-1v-3.65a1 1 0 0 0-.22-.62l-3.48-4.35A1 1 0 0 0 17.52 8H14"/><circle cx="17" cy="18" r="2"/><circle cx="7" cy="18" r="2"/>',
    "flag": '<path d="M4 15s1-1 4-1 5 2 8 2 4-1 4-1V3s-1 1-4 1-5-2-8-2-4 1-4 1z"/><line x1="4" x2="4" y1="22" y2="15"/>',
    "calendar": '<rect width="18" height="18" x="3" y="4" rx="2"/><path d="M16 2v4"/><path d="M8 2v4"/><path d="M3 10h18"/>',
    "sliders": '<line x1="4" x2="20" y1="7" y2="7"/><line x1="4" x2="20" y1="17" y2="17"/><circle cx="9" cy="7" r="2.5"/><circle cx="15" cy="17" r="2.5"/>',
    "building": '<rect width="16" height="20" x="4" y="2" rx="2"/><path d="M9 22v-4h6v4"/><path d="M8 6h.01M16 6h.01M12 6h.01M12 10h.01M12 14h.01M16 10h.01M16 14h.01M8 10h.01M8 14h.01"/>',
}


def icon_svg(name):
    return (
        "<svg viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' "
        f"stroke-linecap='round' stroke-linejoin='round'>{_ICON_PATHS[name]}</svg>"
    )


def sparkline_svg(values, color):
    """Small trend line for a KPI card: area wash + 2px line + end dot on the latest value."""
    vals = [float(v) for v in values]
    if len(vals) < 2:
        return ""
    w, h, pad = 200, 34, 4
    lo, hi = min(vals), max(vals)
    span = (hi - lo) or 1.0
    step = (w - 2 * pad) / (len(vals) - 1)
    pts = [(pad + i * step, pad + (h - 2 * pad) * (1 - (v - lo) / span)) for i, v in enumerate(vals)]
    line = " ".join(f"{x:.1f},{y:.1f}" for x, y in pts)
    area = f"{pts[0][0]:.1f},{h} {line} {pts[-1][0]:.1f},{h}"
    ex, ey = pts[-1]
    return (
        f"<svg viewBox='0 0 {w} {h}' preserveAspectRatio='none'>"
        f"<polygon points='{area}' fill='{color}' fill-opacity='0.10'/>"
        f"<polyline points='{line}' fill='none' stroke='{color}' stroke-width='2' "
        f"stroke-linejoin='round' stroke-linecap='round' vector-effect='non-scaling-stroke'/>"
        f"<circle cx='{ex:.1f}' cy='{ey:.1f}' r='3.5' fill='{color}' stroke='{COLOR_CARD}' stroke-width='2' "
        f"vector-effect='non-scaling-stroke'/>"
        f"</svg>"
    )


def kpi_card(container, icon, accent, label, value, unit=None, delta_text=None, delta_tone=None,
             vs_text=None, spark=None, meter=None):
    """Stat tile. delta_tone: 'good' | 'bad' | 'neutral'. spark: list of values. meter: 0..1 fill."""
    if delta_text:
        tone_color = {"good": COLOR_OK, "bad": COLOR_CRIT}.get(delta_tone, COLOR_MUTED)
        tone_bg = hex_to_rgba(tone_color, 0.12) if delta_tone in ("good", "bad") else hex_to_rgba(COLOR_MUTED, 0.12)
        delta_html = f"<span class='kpi-delta' style='background:{tone_bg}; color:{tone_color};'>{delta_text}</span>"
    else:
        delta_html = ""
    vs_html = f"<span class='kpi-vs'>{vs_text}</span>" if vs_text else ""
    foot_html = f"<div class='kpi-foot'>{delta_html}{vs_html}</div>" if (delta_html or vs_html) else ""
    unit_html = f"<span class='kpi-unit'>{unit}</span>" if unit else ""
    extra_html = ""
    if spark:
        extra_html = f"<div class='kpi-spark'>{sparkline_svg(spark, accent)}</div>"
    elif meter is not None:
        pct = max(0.0, min(1.0, float(meter))) * 100
        extra_html = (
            f"<div class='kpi-meter'><div class='kpi-meter-track'>"
            f"<div class='kpi-meter-fill' style='width:{pct:.1f}%;'></div></div></div>"
        )
    container.markdown(
        f"<div class='kpi-card' style='--accent:{accent};'>"
        f"<div class='kpi-head'><span class='kpi-label'>{label}</span>"
        f"<span class='kpi-icon'>{icon_svg(icon)}</span></div>"
        f"<div class='kpi-value'>{value}{unit_html}</div>"
        f"{foot_html}{extra_html}"
        f"</div>",
        unsafe_allow_html=True,
    )


def card_header(title, subtitle=None, legend=None):
    """Title row printed at the top of a bordered chart container. legend: [(label, color), ...]."""
    html = f"<div class='card-title'>{title}</div>"
    if subtitle:
        html += f"<div class='card-sub'>{subtitle}</div>"
    if legend:
        items = "".join(
            f"<span class='legend-item'><span class='legend-swatch' style='background:{c};'></span>{lbl}</span>"
            for lbl, c in legend
        )
        html += f"<div class='legend-row'>{items}</div>"
    st.markdown(html, unsafe_allow_html=True)


def style_fig(fig, height, y_title=None, x_title=None, show_legend=False):
    """One consistent, quiet chart look: recessive hairline grid, muted axes, card-coloured surface."""
    axis_font = dict(size=12, color=CHART_TEXT, family="Inter, sans-serif")
    fig.update_layout(
        plot_bgcolor=CHART_BG, paper_bgcolor=CHART_BG,
        font=dict(family="Inter, sans-serif", color=CHART_TEXT, size=12),
        height=height, margin=dict(l=4, r=12, t=8, b=4),
        showlegend=show_legend,
        legend=dict(orientation="h", y=1.02, yanchor="bottom", x=0, font=dict(size=12, color=CHART_TEXT)),
        hoverlabel=dict(bgcolor=COLOR_CARD, bordercolor=COLOR_BORDER, font=dict(family="Inter, sans-serif", size=13, color=COLOR_TEXT)),
        bargap=0.35,
    )
    fig.update_xaxes(
        showgrid=False, zeroline=False, linecolor=CHART_AXIS, linewidth=1, ticks="",
        tickfont=axis_font, title=x_title, title_font=axis_font,
    )
    fig.update_yaxes(
        gridcolor=CHART_GRID, gridwidth=1, zeroline=False, linecolor=CHART_AXIS, ticks="",
        tickfont=axis_font, title=y_title, title_font=axis_font,
        exponentformat="none", separatethousands=True,
    )
    return fig


PLOTLY_CONFIG = {"displayModeBar": False}


# ---------------------------------------------------------------------------
# Light / dark toggle — a small sun / moon tile in the sidebar's top row,
# beside the collapse arrow. Streamlit has no Python API to change the viewer's theme, so the
# click runs a tiny script that picks Light/Dark in Streamlit's own ⋮ menu
# (which also remembers the choice in the browser) and then presses a hidden
# button to rerun the app, so our cards and charts redraw in the new palette.
# ---------------------------------------------------------------------------
st.markdown(f"""
<style>
[data-testid="stSidebarContent"] {{ position: relative; }}
.st-key-theme_toggle {{
    position: absolute; top: 0.85rem; left: 1.4rem; z-index: 999991; width: auto !important;
}}
.st-key-theme_toggle button {{
    width: 36px !important; height: 36px !important; min-height: 36px !important; padding: 0 !important;
    border-radius: 11px !important; font-size: 19px !important; line-height: 1 !important;
    background: {"#1C2740" if _IS_DARK else "#EEF0F7"} !important;
    border: 1px solid {COLOR_BORDER} !important; box-shadow: 0 1px 2px {SHADOW_SM} !important;
}}
.st-key-theme_toggle button p {{ font-size: 19px !important; line-height: 1 !important; color: inherit !important; }}
.st-key-theme_toggle button:hover {{
    background: {"#24314F" if _IS_DARK else "#E2E6F2"} !important; transform: translateY(-1px);
}}
.st-key-theme_sync {{ display: none !important; }}
[data-testid="stElementContainer"]:has(iframe[height="0"]) {{ display: none !important; }}
</style>
""", unsafe_allow_html=True)

_target_theme = "Light" if _IS_DARK else "Dark"
with st.sidebar, st.container(key="theme_toggle"):
    _toggle_clicked = st.button(
        "☀️" if _IS_DARK else "🌙", key="theme_toggle_btn",
        help=f"Switch to {_target_theme.lower()} mode",
    )
with st.container(key="theme_sync"):
    st.button("sync theme", key="theme_sync_btn")

if _toggle_clicked:
    components.html(f"""
<script>
(function () {{
  const d = window.parent.document;
  const target = "{_target_theme}";
  function rerun() {{
    const sync = d.querySelector('.st-key-theme_sync button');
    if (sync) sync.click();
  }}
  function pick(attempt) {{
    const item = d.querySelector('[data-testid="stMainMenuItem-theme-' + target + '"]');
    if (item) {{
      item.click();
      d.body.dispatchEvent(new KeyboardEvent('keydown', {{key: 'Escape', bubbles: true}}));
      setTimeout(rerun, 150);
    }} else if (attempt < 20) {{
      setTimeout(function () {{ pick(attempt + 1); }}, 50);
    }}
  }}
  const menu = d.querySelector('[data-testid="stMainMenu"] button') || d.querySelector('[data-testid="stMainMenu"]');
  if (menu) {{ menu.click(); pick(0); }}
}})();
</script>
""", height=0)



# Permanent company branding — edit this constant directly to change the
# name. To change the logo, replace the file at .streamlit/logo.png.
COMPANY_NAME = "Supreme Facility Management Limited"


# ---------------------------------------------------------------------------
# Login credentials + company branding (stored on disk, editable at runtime)
# ---------------------------------------------------------------------------
CREDENTIALS_PATH = Path(__file__).parent / ".streamlit" / "credentials.json"
BRANDING_PATH = Path(__file__).parent / ".streamlit" / "branding.json"
LOGO_PATH = Path(__file__).parent / ".streamlit" / "logo.png"


def load_credentials():
    # On Streamlit Community Cloud, prefer the Secrets manager (configured in the
    # app's dashboard, never stored in the git repo) over the local file.
    try:
        if "ADMIN_USERNAME" in st.secrets and "ADMIN_PASSWORD" in st.secrets:
            return {"username": st.secrets["ADMIN_USERNAME"], "password": st.secrets["ADMIN_PASSWORD"]}
    except Exception:
        pass
    if CREDENTIALS_PATH.exists():
        try:
            return json.loads(CREDENTIALS_PATH.read_text())
        except Exception:
            return None
    return None


def using_cloud_secrets():
    try:
        return "ADMIN_USERNAME" in st.secrets and "ADMIN_PASSWORD" in st.secrets
    except Exception:
        return False


def save_credentials(creds):
    CREDENTIALS_PATH.parent.mkdir(exist_ok=True)
    CREDENTIALS_PATH.write_text(json.dumps(creds, indent=2))


def load_branding():
    if BRANDING_PATH.exists():
        try:
            return json.loads(BRANDING_PATH.read_text())
        except Exception:
            return {}
    return {}


def save_branding(data):
    BRANDING_PATH.parent.mkdir(exist_ok=True)
    BRANDING_PATH.write_text(json.dumps(data, indent=2))


# ---------------------------------------------------------------------------
# Brand header band (shown at the top of every page, and on the login screen)
# ---------------------------------------------------------------------------
def render_brand_header(title="GPS vs MIS Dashboard", subtitle=None):
    def chip(icon, text):
        return f"<span class='hero-chip'>{icon_svg(icon)}{text}</span>"

    chips = ""
    if st.session_state.get("authenticated"):
        if st.session_state.get("month_label"):
            chips += chip("calendar", st.session_state.month_label)
        if st.session_state.get("threshold") is not None:
            chips += chip("sliders", f"Flag limit ±{st.session_state.threshold}%")
    sub_html = f"<div class='hero-sub'>{subtitle}</div>" if subtitle else ""
    st.markdown(
        f"<div class='brand-hero'>"
        f"<div style='position:relative; z-index:1;'>"
        f"<div class='hero-eyebrow'>{COMPANY_NAME} · Fleet telemetry reconciliation</div>"
        f"<div class='hero-title'>{title}</div>{sub_html}</div>"
        f"<div class='hero-chips'>{chips}</div>"
        f"</div>",
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Login
# ---------------------------------------------------------------------------
def check_login():
    if st.session_state.get("authenticated"):
        return True

    render_brand_header("GPS vs MIS Fleet Dashboard", "Sign in with your admin account to view this month's reconciliation.")
    _, mid, _ = st.columns([1, 1.3, 1])
    with mid:
        st.markdown("#### Admin login")
        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Log in", use_container_width=True)
        st.caption("🔒 Confidential — for authorised Supreme Facility Management staff only.")

    if submitted:
        creds = load_credentials()
        if not creds:
            st.error(
                "No admin credentials configured yet. Run start.bat again, or copy "
                "`.streamlit/credentials.example.json` to `.streamlit/credentials.json` "
                "and set your own username/password (see README.md)."
            )
            return False
        if username == creds.get("username") and password == creds.get("password"):
            st.session_state.authenticated = True
            st.session_state.current_username = username
            st.rerun()
        else:
            st.error("Incorrect username or password.")
    return False


if not check_login():
    st.stop()

# Streamlit always renders the page navigation above st.sidebar content, so
# the logo and company name are injected at the top of the navigation block
# via CSS: the logo (embedded as a data URI) sits on a white tile above the name.
_company_css = COMPANY_NAME.replace("\\", "\\\\").replace('"', '\\"')
_logo_css = ""
if LOGO_PATH.exists():
    _logo_b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode()
    _logo_css = (
        f"background: #FFFFFF url('data:image/png;base64,{_logo_b64}') no-repeat center 10px / auto 80px;"
        "padding-top: 96px !important; border-radius: 12px; border: 1px solid rgba(128,128,128,0.18);"
        "padding-bottom: 10px !important; color: #0F172A !important;"
    )
st.markdown(
    f"""<style>
    [data-testid="stSidebarNav"]::before {{
        content: "{_company_css}";
        display: block;
        text-align: center;
        font-size: 0.95rem;
        font-weight: 700;
        line-height: 1.3;
        padding: 0 0.6rem 0.8rem;
        margin-bottom: 0.8rem;
        {_logo_css}
    }}
    </style>""",
    unsafe_allow_html=True,
)

with st.sidebar:
    st.caption(f"Logged in as **{st.session_state.get('current_username', 'admin')}**")
    if st.button("Log out"):
        st.session_state.authenticated = False
        st.rerun()

    with st.expander("🔑 Change password"):
        if using_cloud_secrets():
            st.info(
                "This app is running on Streamlit Community Cloud, where the password "
                "comes from the app's **Secrets** settings, not this file. To change it: "
                "go to your app on share.streamlit.io → Settings → Secrets, update "
                "ADMIN_PASSWORD there, and save (the app restarts automatically)."
            )
        else:
            with st.form("change_password_form"):
                old_pw = st.text_input("Current password", type="password")
                new_pw = st.text_input("New password", type="password")
                confirm_pw = st.text_input("Confirm new password", type="password")
                pw_submit = st.form_submit_button("Update password")
            if pw_submit:
                creds = load_credentials() or {}
                if old_pw != creds.get("password"):
                    st.error("Current password is incorrect.")
                elif not new_pw:
                    st.error("New password can't be empty.")
                elif new_pw != confirm_pw:
                    st.error("New password and confirmation don't match.")
                else:
                    creds["password"] = new_pw
                    save_credentials(creds)
                    st.success("Password updated — use it next time you log in.")

# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------
def read_sheet(ws):
    day_cols = []
    for c in range(4, ws.max_column + 1):
        h = ws.cell(row=4, column=c).value
        if isinstance(h, (int, float)):
            day_cols.append(c)
        elif h is None and day_cols:
            break
    rows = {}
    for r in range(5, ws.max_row + 1):
        veh = ws.cell(row=r, column=2).value
        if not veh:
            continue
        cc = ws.cell(row=r, column=3).value
        vals = []
        for c in day_cols:
            v = ws.cell(row=r, column=c).value
            vals.append(v if isinstance(v, (int, float)) else 0)
        rows[str(veh).strip()] = {"cost_center": (cc or "Unassigned"), "days": vals}
    return rows


@st.cache_data(show_spinner="Reading workbook...")
def build_dataset(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    gps_name = next((n for n in wb.sheetnames if n.strip().upper() == "GPS"), None)
    mis_name = next((n for n in wb.sheetnames if n.strip().upper() == "MIS"), None)
    if not gps_name or not mis_name:
        raise ValueError('Could not find sheets named "GPS" and "MIS" in this file.')

    gps_rows = read_sheet(wb[gps_name])
    mis_rows = read_sheet(wb[mis_name])
    n_days = max(
        max((len(v["days"]) for v in gps_rows.values()), default=0),
        max((len(v["days"]) for v in mis_rows.values()), default=0),
    )
    all_veh = list(dict.fromkeys(list(gps_rows.keys()) + list(mis_rows.keys())))

    records = []
    daily_records = []
    for veh in all_veh:
        g = gps_rows.get(veh)
        m = mis_rows.get(veh)
        cc = (g or m)["cost_center"]
        gdays = (g["days"] if g else []) + [0] * n_days
        mdays = (m["days"] if m else []) + [0] * n_days
        gdays, mdays = gdays[:n_days], mdays[:n_days]
        tg, tm = sum(gdays), sum(mdays)
        td = tm - tg
        dpct = (td / tg) if tg else 0.0
        source = "Both" if (g and m) else ("GPS only" if g else "MIS only")
        records.append({
            "Vehicle": veh, "Site": cc, "Total GPS": round(tg, 1), "Total MIS": round(tm, 1),
            "Diff": round(td, 1), "Diff %": dpct, "Source": source,
        })
        for d in range(n_days):
            daily_records.append({"Vehicle": veh, "Day": d + 1, "GPS": gdays[d], "MIS": mdays[d]})

    return pd.DataFrame(records), pd.DataFrame(daily_records), n_days


def classify(row, threshold_pct):
    t = threshold_pct / 100
    tg, tm, dpct = row["Total GPS"], row["Total MIS"], row["Diff %"]
    if tg == 0 and tm > 0:
        return "Check GPS device", "critical", "No GPS km recorded all month while MIS shows movement. Confirm the device is powered, fitted, and reporting."
    if tm == 0 and tg > 0:
        return "File missing MIS log", "critical", "GPS shows movement but no MIS entries were filed. Follow up with the site log-keeper."
    if tg == 0 and tm == 0:
        return "No data either side", "warn", "Both sources show zero for the month. Confirm the vehicle was actually in service."
    if abs(dpct) > t:
        if dpct > 0:
            return "Audit MIS entries", "critical", f"MIS log shows {abs(dpct)*100:.1f}% more km than GPS. Cross-check trip sheets and fuel/DA claims for over-reporting."
        return "Verify unrecorded trips", "critical", f"GPS shows {abs(dpct)*100:.1f}% more km than MIS. Check for trips run but not logged."
    if abs(dpct) > t * 0.5:
        return "Keep an eye on it", "watch", f"Diff is under the {threshold_pct}% threshold but trending up. Worth a spot-check next month."
    return "No action needed", "ok", "GPS and MIS agree within tolerance."


# ---------------------------------------------------------------------------
# Month-over-month history (saved to disk, survives restarts, never
# overwritten by a new upload unless it's the same month being re-saved)
# ---------------------------------------------------------------------------
HISTORY_PATH = Path(__file__).parent / "history.json"

MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


def load_history():
    if HISTORY_PATH.exists():
        try:
            return json.loads(HISTORY_PATH.read_text())
        except Exception:
            return {}
    return {}


def save_history(hist):
    HISTORY_PATH.write_text(json.dumps(hist, indent=2))


def guess_month_from_filename(filename):
    m = re.search(r"(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\D{0,5}(20\d{2})", filename, re.IGNORECASE)
    if m:
        mon = MONTH_MAP[m.group(1).lower()]
        year = int(m.group(2))
        key = f"{year:04d}-{mon:02d}"
        label = datetime(year, mon, 1).strftime("%B %Y")
        return key, label
    now = datetime.now()
    return now.strftime("%Y-%m"), now.strftime("%B %Y")


def get_previous_entry(hist, current_key):
    keys = sorted(k for k in hist.keys() if k < current_key)
    return hist[keys[-1]] if keys else None


# ---------------------------------------------------------------------------
# Data source — upload / pick a saved month / set the flag threshold.
# Lives in the sidebar so it's available on every page without re-uploading.
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("---")
    st.markdown("#### Data")
    uploaded = st.file_uploader("Drop this month's file here (needs GPS and MIS sheets)", type=["xlsx", "xls"])

hist = load_history()
sorted_hist_entries = sorted(hist.values(), key=lambda x: x["key"], reverse=True)
month_options = [e["label"] for e in sorted_hist_entries]

if uploaded is None and not month_options:
    render_brand_header()
    st.info("Upload your monthly GPS + MIS workbook using the sidebar to see the analysis.")
    st.stop()

with st.sidebar:
    view_options = (["📤 Uploaded file"] if uploaded is not None else []) + month_options
    view_choice = st.selectbox("Viewing", view_options, index=0)

    has_daily = False

    if view_choice == "📤 Uploaded file":
        try:
            df, daily_df, n_days = build_dataset(uploaded.getvalue())
        except Exception as e:
            st.error(str(e))
            st.stop()
        has_daily = True

        guessed_key, guessed_label = guess_month_from_filename(uploaded.name)
        month_label = st.text_input("Month label for this file (edit if the guess is wrong)", value=guessed_label)
        try:
            month_key = datetime.strptime(month_label.strip(), "%B %Y").strftime("%Y-%m")
        except ValueError:
            month_key = guessed_key
    else:
        entry = hist[next(e["key"] for e in sorted_hist_entries if e["label"] == view_choice)]
        df = pd.DataFrame(entry["vehicle_totals"])
        df["Diff"] = df["Total MIS"] - df["Total GPS"]
        df["Diff %"] = df.apply(lambda r: (r["Diff"] / r["Total GPS"]) if r["Total GPS"] else 0, axis=1)
        df["Source"] = "Both"
        month_key = entry["key"]

        daily_compact = entry.get("daily")
        if daily_compact:
            has_daily = True
            daily_rows = []
            for veh, series in daily_compact.items():
                gvals, mvals = series["g"], series["m"]
                for d in range(len(gvals)):
                    daily_rows.append({"Vehicle": veh, "Day": d + 1, "GPS": gvals[d], "MIS": mvals[d]})
            daily_df = pd.DataFrame(daily_rows)
            st.caption(f"Viewing saved history for **{view_choice}** — including day-wise breakdown.")
        else:
            daily_df = pd.DataFrame(columns=["Vehicle", "Day", "GPS", "MIS"])
            st.caption(f"Viewing saved history for **{view_choice}** — this month was saved before day-wise data was stored, so only totals are available.")

    threshold = st.slider("Flag threshold (%)  — flag vehicles where |Diff %| exceeds this", 5, 100, 20, step=5)

action_info = df.apply(lambda r: classify(r, threshold), axis=1)
df["Action"] = action_info.apply(lambda x: x[0])
df["Severity"] = action_info.apply(lambda x: x[1])
df["ActionDetail"] = action_info.apply(lambda x: x[2])

total_gps = df["Total GPS"].sum()
total_mis = df["Total MIS"].sum()
flagged = (df["Action"] != "No action needed").sum()

if view_choice == "📤 Uploaded file":
    daily_compact = {}
    for veh, grp in daily_df.groupby("Vehicle"):
        grp_sorted = grp.sort_values("Day")
        daily_compact[veh] = {
            "g": [round(float(v), 1) for v in grp_sorted["GPS"].tolist()],
            "m": [round(float(v), 1) for v in grp_sorted["MIS"].tolist()],
        }
    hist[month_key] = {
        "key": month_key, "label": month_label.strip(), "saved_at": datetime.now().isoformat(),
        "total_gps": float(total_gps), "total_mis": float(total_mis),
        "vehicles": int(len(df)), "sites": int(df["Site"].nunique()), "flagged": int(flagged),
        "site_summary": df.groupby("Site", as_index=False).agg(
            Total_GPS=("Total GPS", "sum"), Total_MIS=("Total MIS", "sum")
        ).to_dict("records"),
        "vehicle_totals": df[["Vehicle", "Site", "Total GPS", "Total MIS"]].to_dict("records"),
        "daily": daily_compact,
    }
    save_history(hist)
    with st.sidebar:
        st.caption(f"✓ Autosaved as **{month_label.strip()}** — {len(hist)} month(s) in history now.")

prev_entry = get_previous_entry(hist, month_key)

with st.sidebar:
    st.caption(f"Generated {datetime.now().strftime('%d %b %Y, %H:%M')} — runs entirely on your machine, nothing uploaded externally.")

# Make the prepared data available to every page.
st.session_state.df = df
st.session_state.daily_df = daily_df
st.session_state.has_daily = has_daily
st.session_state.threshold = threshold
st.session_state.prev_entry = prev_entry
st.session_state.month_key = month_key
st.session_state.month_label = month_label.strip() if view_choice == "📤 Uploaded file" else view_choice


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------
SEVERITY_ORDER = ["critical", "warn", "watch", "ok"]


def history_up_to(month_key, limit=6):
    """Saved months up to and including month_key, oldest first (for KPI sparklines)."""
    entries = sorted(load_history().values(), key=lambda e: e["key"])
    return [e for e in entries if e["key"] <= month_key][-limit:]


def daily_trend_fig(data, height=320, value_fmt=",.0f"):
    """GPS vs MIS km per day of month — two 2px lines, unified hover."""
    fig = go.Figure()
    for col, name, color in (("GPS", "GPS", COLOR_GPS), ("MIS", "MIS", COLOR_MIS)):
        fig.add_trace(go.Scatter(
            x=data["Day"], y=data[col], name=name, mode="lines+markers",
            line=dict(color=color, width=2, shape="linear"),
            marker=dict(size=7, color=color, line=dict(color=CHART_BG, width=2)),
            hovertemplate=f"{name}: %{{y:{value_fmt}}} km<extra></extra>",
        ))
    style_fig(fig, height, y_title="Km", x_title="Day of month")
    fig.update_xaxes(dtick=2)
    fig.update_yaxes(tickformat=",.0f", rangemode="tozero")
    fig.update_layout(hovermode="x unified")
    return fig


def site_deviation_fig(site_summary, threshold_pct, height):
    """Diverging horizontal bars: MIS vs GPS difference % per site, coloured by status,
    with the ± flag threshold drawn as reference lines. Largest deviation at the top."""
    data = site_summary.copy()
    data["pct"] = data["Diff %"] * 100
    data = data.reindex(data["pct"].abs().sort_values(ascending=True).index)
    colors = data["Severity"].map(STATUS_FILL)
    fig = go.Figure(go.Bar(
        y=data["Site"], x=data["pct"], orientation="h",
        marker=dict(color=colors, cornerradius=4),
        text=data["pct"].apply(lambda v: f"{v:+.1f}%"), textposition="outside", cliponaxis=False,
        textfont=dict(size=12, color=COLOR_TEXT, family="Inter, sans-serif"),
        customdata=list(zip(data["Total_GPS"], data["Total_MIS"], data["Severity"].map(STATUS_LABEL))),
        hovertemplate="<b>%{y}</b><br>Diff: %{x:+.1f}%<br>GPS: %{customdata[0]:,.0f} km"
                      "<br>MIS: %{customdata[1]:,.0f} km<br>Status: %{customdata[2]}<extra></extra>",
    ))
    style_fig(fig, height, x_title="MIS vs GPS difference (%)")
    t = float(threshold_pct)
    lo = min(data["pct"].min(), -t) if len(data) else -t
    hi = max(data["pct"].max(), t) if len(data) else t
    pad = (hi - lo) * 0.18
    fig.update_xaxes(range=[lo - pad, hi + pad], showgrid=True, gridcolor=CHART_GRID, ticksuffix="%")
    fig.update_yaxes(showgrid=False, automargin=True, tickfont=dict(size=12.5, color=COLOR_TEXT))
    fig.add_vline(x=0, line=dict(color=CHART_AXIS, width=1))
    for x in (-t, t):
        fig.add_vline(x=x, line=dict(color=COLOR_MUTED, width=1, dash="dot"))
    fig.add_annotation(x=t, y=1, yref="paper", yanchor="bottom", text=f"+{t:.0f}% limit", showarrow=False,
                       font=dict(size=11, color=COLOR_MUTED))
    fig.add_annotation(x=-t, y=1, yref="paper", yanchor="bottom", text=f"−{t:.0f}% limit", showarrow=False,
                       font=dict(size=11, color=COLOR_MUTED))
    fig.update_layout(margin=dict(l=4, r=16, t=24, b=4), bargap=0.32)
    return fig


def fleet_status_fig(df):
    """Donut of vehicles by status, flagged count in the centre."""
    counts = df["Severity"].value_counts()
    sev = [s for s in SEVERITY_ORDER if counts.get(s, 0) > 0]
    vals = [int(counts[s]) for s in sev]
    flagged = int((df["Action"] != "No action needed").sum())
    fig = go.Figure(go.Pie(
        labels=[STATUS_LABEL[s] for s in sev], values=vals, hole=0.68, sort=False, direction="clockwise",
        marker=dict(colors=[STATUS_FILL[s] for s in sev], line=dict(color=CHART_BG, width=3)),
        textinfo="none",
        hovertemplate="<b>%{label}</b><br>%{value} vehicles (%{percent})<extra></extra>",
    ))
    fig.update_layout(
        plot_bgcolor=CHART_BG, paper_bgcolor=CHART_BG, height=250, margin=dict(l=0, r=0, t=6, b=6),
        showlegend=False, font=dict(family="Inter, sans-serif"),
        hoverlabel=dict(bgcolor=COLOR_CARD, bordercolor=COLOR_BORDER, font=dict(family="Inter, sans-serif", size=13, color=COLOR_TEXT)),
        annotations=[
            dict(text=f"<b>{flagged}</b>", x=0.5, y=0.55, showarrow=False, font=dict(size=30, color=COLOR_TEXT)),
            dict(text="flagged", x=0.5, y=0.38, showarrow=False, font=dict(size=13, color=COLOR_MUTED)),
        ],
    )
    return fig, [(f"{STATUS_LABEL[s]} · {int(counts[s])}", STATUS_FILL[s]) for s in sev]


def monthly_totals_fig(entries, height=300):
    """Grouped columns: total GPS vs MIS km per saved month."""
    months = [e["label"] for e in entries]
    fig = go.Figure()
    for key, name, color in (("total_gps", "GPS", COLOR_GPS), ("total_mis", "MIS", COLOR_MIS)):
        fig.add_trace(go.Bar(
            x=months, y=[e[key] for e in entries], name=name,
            marker=dict(color=color, cornerradius=4),
            hovertemplate=f"<b>%{{x}}</b><br>{name}: %{{y:,.0f}} km<extra></extra>",
        ))
    style_fig(fig, height, y_title="Km")
    fig.update_yaxes(tickformat=",.0f")
    fig.update_layout(barmode="group", bargap=0.38, bargroupgap=0.12)
    return fig


def monthly_diff_fig(entries, threshold_pct, height=240):
    """Overall MIS vs GPS difference % per month, coloured by the same status rule as sites."""
    t = threshold_pct / 100
    months, pcts, colors, labels = [], [], [], []
    for e in entries:
        p = ((e["total_mis"] - e["total_gps"]) / e["total_gps"]) if e["total_gps"] else 0
        sev = "critical" if abs(p) > t else ("watch" if abs(p) > t * 0.5 else "ok")
        months.append(e["label"])
        pcts.append(p * 100)
        colors.append(STATUS_FILL[sev])
        labels.append(STATUS_LABEL[sev])
    fig = go.Figure(go.Bar(
        x=months, y=pcts, marker=dict(color=colors, cornerradius=4),
        text=[f"{v:+.1f}%" for v in pcts], textposition="outside", cliponaxis=False,
        textfont=dict(size=12, color=COLOR_TEXT), customdata=labels,
        hovertemplate="<b>%{x}</b><br>Diff: %{y:+.1f}%<br>Status: %{customdata}<extra></extra>",
    ))
    style_fig(fig, height, y_title="Diff %")
    fig.update_yaxes(ticksuffix="%")
    fig.add_hline(y=0, line=dict(color=CHART_AXIS, width=1))
    for y in (-threshold_pct, threshold_pct):
        fig.add_hline(y=y, line=dict(color=COLOR_MUTED, width=1, dash="dot"))
    fig.update_layout(bargap=0.55, margin=dict(l=4, r=12, t=24, b=4))
    return fig


def build_insights(df, prev_entry, threshold):
    """Plain-language findings for the month, most important first.
    Returns (match_rate 0..1, [(tone, html_text), ...]) where tone is good | warn | bad."""
    n = len(df)
    match_rate = (df["Severity"] == "ok").sum() / n if n else 0.0
    flagged_df = df[df["Action"] != "No action needed"]
    total_gps, total_mis = df["Total GPS"].sum(), df["Total MIS"].sum()
    net = total_mis - total_gps
    pct = (net / total_gps) if total_gps else 0.0
    items = []

    if abs(pct) * 100 > threshold:
        side = "less" if net < 0 else "more"
        items.append(("bad", f"MIS logs record <b>{abs(net):,.0f} km {side}</b> than GPS across the fleet "
                             f"({pct*100:+.1f}%) — outside the ±{threshold}% limit."))
    else:
        items.append(("good", f"Fleet-wide, MIS and GPS agree within the ±{threshold}% limit ({pct*100:+.1f}%)."))

    missing_mis = df[df["Action"] == "File missing MIS log"]
    if len(missing_mis):
        km = missing_mis["Total GPS"].sum()
        share = (km / total_gps * 100) if total_gps else 0
        items.append(("bad", f"<b>{len(missing_mis)} vehicles</b> moved <b>{km:,.0f} km</b> on GPS with no MIS entry "
                             f"({share:.0f}% of all GPS km) — follow up with site log-keepers."))

    no_gps = df[df["Action"] == "Check GPS device"]
    if len(no_gps):
        items.append(("warn", f"<b>{len(no_gps)} vehicles</b> show MIS km but no GPS signal — check the devices are fitted and reporting."))

    over_report = df[df["Action"] == "Audit MIS entries"]
    if len(over_report):
        items.append(("warn", f"<b>{len(over_report)} vehicles</b> have MIS km well above GPS — audit trip sheets and claims."))

    if len(flagged_df):
        by_site = flagged_df.groupby("Site").size().sort_values(ascending=False)
        top_site, top_n = by_site.index[0], int(by_site.iloc[0])
        site_total = int((df["Site"] == top_site).sum())
        items.append(("warn", f"<b>{top_site}</b> needs the most attention: {top_n} of its {site_total} vehicles are flagged."))

    if prev_entry and isinstance(prev_entry.get("flagged"), (int, float)):
        before, now = int(prev_entry["flagged"]), len(flagged_df)
        if now != before:
            tone = "bad" if now > before else "good"
            items.append((tone, f"Flagged vehicles went <b>{'up' if now > before else 'down'} from {before} to {now}</b> "
                                f"compared with {prev_entry['label']}."))
    return match_rate, items[:6]


def render_insights(df, prev_entry, threshold, month_label):
    """Reconciliation score ring + key findings list, shown on the Overview page."""
    match_rate, items = build_insights(df, prev_entry, threshold)
    n = len(df)
    ok = int((df["Severity"] == "ok").sum())
    ring = COLOR_OK if match_rate >= 0.8 else (COLOR_WARN if match_rate >= 0.6 else COLOR_CRIT)
    tone_color = {"good": COLOR_OK, "warn": COLOR_WARN, "bad": COLOR_CRIT}
    rows = "".join(
        f"<div class='insight-row'><span class='insight-dot' style='background:{tone_color[t]};'></span>"
        f"<span class='insight-text'>{txt}</span></div>"
        for t, txt in items
    )
    st.markdown(
        f"<div class='insight-card'>"
        f"<div class='score'><div class='score-ring' style='--p:{match_rate*100:.1f}; --c:{ring};'>"
        f"<div class='score-inner'><div class='score-val'>{match_rate*100:.0f}%</div><div class='score-lbl'>matched</div></div></div>"
        f"<div class='score-cap'>Reconciliation score</div>"
        f"<div class='score-sub'>{ok} of {n} vehicles agree within tolerance</div></div>"
        f"<div class='insight-body'><div class='insight-title'>Key findings — {month_label}</div>{rows}</div>"
        f"</div>",
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def build_excel_report(df, month_label, threshold, generated_at):
    """Management-ready workbook: Summary (+ key findings), Action plan, Sites, All vehicles."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    site = compute_site_summary(df, threshold)
    total_gps, total_mis = df["Total GPS"].sum(), df["Total MIS"].sum()
    match_rate, findings = build_insights(df, None, threshold)
    flagged = df[df["Action"] != "No action needed"].copy()

    summary = pd.DataFrame([
        ("Company", COMPANY_NAME), ("Report", "GPS vs MIS fleet reconciliation"), ("Month", month_label),
        ("Flag limit", f"±{threshold}%"), ("Total GPS km", round(total_gps, 1)), ("Total MIS km", round(total_mis, 1)),
        ("Net difference (MIS − GPS) km", round(total_mis - total_gps, 1)),
        ("Net difference %", ((total_mis - total_gps) / total_gps) if total_gps else 0),
        ("Vehicles", len(df)), ("Sites", df["Site"].nunique()), ("Flagged vehicles", len(flagged)),
        ("Reconciliation score (vehicles matched)", match_rate), ("Generated", generated_at),
    ], columns=["Item", "Value"])

    sev_rank = {s: i for i, s in enumerate(SEVERITY_ORDER)}
    flagged["_r"] = flagged["Severity"].map(sev_rank)
    flagged["_a"] = flagged["Diff %"].abs()
    action_plan = flagged.sort_values(["_r", "_a"], ascending=[True, False])[
        ["Vehicle", "Site", "Total GPS", "Total MIS", "Diff", "Diff %", "Action", "ActionDetail"]
    ].rename(columns={"Diff": "Diff km", "ActionDetail": "What to do"})
    sites = site.assign(Status=site["Severity"].map(STATUS_LABEL)).drop(columns=["Severity"]).rename(
        columns={"Total_GPS": "Total GPS", "Total_MIS": "Total MIS"})
    vehicles = df.sort_values("Diff %", key=abs, ascending=False)[
        ["Vehicle", "Site", "Total GPS", "Total MIS", "Diff", "Diff %", "Action", "Source"]
    ].rename(columns={"Diff": "Diff km"})

    header_fill = PatternFill("solid", fgColor="13317A")
    header_font = Font(bold=True, color="FFFFFF")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, frame in (("Summary", summary), ("Action plan", action_plan), ("Sites", sites), ("All vehicles", vehicles)):
            frame.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.fill, cell.font = header_fill, header_font
                cell.alignment = Alignment(vertical="center")
            for idx, col in enumerate(frame.columns, start=1):
                letter = get_column_letter(idx)
                width = max([len(str(col))] + [len(str(v)) for v in frame[col].head(300)]) + 2
                ws.column_dimensions[letter].width = min(width, 70)
                if name == "Summary":
                    continue
                fmt = "0.0%" if "%" in str(col) else ("#,##0.0" if ("km" in str(col) or "Total" in str(col)) else None)
                if fmt:
                    for cell in ws[letter][1:]:
                        cell.number_format = fmt
        ws = xw.sheets["Summary"]
        for row in ws.iter_rows(min_row=2):
            item = str(row[0].value)
            if item.endswith("%") or item.startswith("Reconciliation"):
                row[1].number_format = "0.0%"
            elif "km" in item:
                row[1].number_format = "#,##0.0"
        start = len(summary) + 3
        ws.cell(row=start, column=1, value="Key findings").font = Font(bold=True, size=12)
        for i, (_, txt) in enumerate(findings, start=1):
            ws.cell(row=start + i, column=1, value="• " + re.sub(r"<[^>]+>", "", txt))
    return buf.getvalue()


def report_download_button(key):
    month_label = st.session_state.month_label
    data = build_excel_report(st.session_state.df, month_label, st.session_state.threshold,
                              datetime.now().strftime("%d %b %Y"))
    st.download_button(
        "⬇  Download Excel report", data=data,
        file_name=f"GPS_vs_MIS_Report_{month_label.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, key=key,
    )


def render_footer():
    st.markdown(
        f"<div class='app-footer'>"
        f"<span><b>{COMPANY_NAME}</b> · GPS vs MIS Fleet Reconciliation</span>"
        f"<span>Confidential — for internal use only · Report for {st.session_state.get('month_label', '—')} · "
        f"Generated {datetime.now().strftime('%d %b %Y, %H:%M')}</span>"
        f"</div>",
        unsafe_allow_html=True,
    )


def page_overview():
    df = st.session_state.df
    daily_df = st.session_state.daily_df
    has_daily = st.session_state.has_daily
    prev_entry = st.session_state.prev_entry
    threshold = st.session_state.threshold

    render_brand_header("Fleet overview", "GPS-tracked vs MIS-logged kilometres, reconciled for every vehicle and site.")

    total_gps = df["Total GPS"].sum()
    total_mis = df["Total MIS"].sum()
    overall_diff = total_mis - total_gps
    overall_pct = (overall_diff / total_gps) if total_gps else 0
    flagged = int((df["Action"] != "No action needed").sum())
    n_sites = df["Site"].nunique()

    past = history_up_to(st.session_state.month_key)
    has_trend = len(past) >= 2

    def km_delta(current, previous):
        d = current - previous
        return f"{'▲' if d >= 0 else '▼'} {abs(d):,.0f} km"

    vs_prev = f"vs {prev_entry['label']}" if prev_entry else None

    c1, c2, c3, c4, c5 = st.columns(5)
    kpi_card(c1, "gps", COLOR_GPS, "Total GPS km", f"{total_gps:,.0f}",
             delta_text=km_delta(total_gps, prev_entry["total_gps"]) if prev_entry else None,
             delta_tone="neutral", vs_text=vs_prev,
             spark=[e["total_gps"] for e in past] if has_trend else None)
    kpi_card(c2, "mis", COLOR_MIS, "Total MIS km", f"{total_mis:,.0f}",
             delta_text=km_delta(total_mis, prev_entry["total_mis"]) if prev_entry else None,
             delta_tone="neutral", vs_text=vs_prev,
             spark=[e["total_mis"] for e in past] if has_trend else None)
    over_limit = abs(overall_pct) * 100 > threshold
    kpi_card(c3, "diff", COLOR_DIFF, "Net difference", f"{overall_pct*100:+.1f}", unit="%",
             delta_text=f"{overall_diff:+,.0f} km", delta_tone="bad" if over_limit else "good",
             vs_text=f"{'Outside' if over_limit else 'Within'} ±{threshold}% limit",
             spark=[((e["total_mis"] - e["total_gps"]) / e["total_gps"] * 100) if e["total_gps"] else 0 for e in past]
             if has_trend else None)
    kpi_card(c4, "truck", COLOR_VEHICLES, "Vehicles", f"{len(df):,}",
             vs_text=f"across {n_sites} sites",
             spark=[e["vehicles"] for e in past] if has_trend else None)
    kpi_card(c5, "flag", COLOR_FLAGGED, "Flagged vehicles", f"{flagged:,}",
             delta_text=f"{flagged/len(df)*100:.0f}% of fleet" if len(df) else None, delta_tone="bad" if flagged else "good",
             meter=(flagged / len(df)) if len(df) else 0)

    if not prev_entry:
        st.caption("No previous month saved yet — month-on-month comparisons appear here once another month is saved.")
    if has_trend:
        st.caption(f"Trend lines in the cards show the last {len(past)} saved months, ending with {past[-1]['label']}.")

    st.write("")
    render_insights(df, prev_entry, threshold, st.session_state.month_label)

    if has_daily:
        with st.container(border=True, key="card_1"):
            card_header("Day-wise kilometres, whole fleet", "Total km recorded each day by GPS and logged in MIS",
                        legend=[("GPS", COLOR_GPS), ("MIS", COLOR_MIS)])
            trend = daily_df.groupby("Day", as_index=False)[["GPS", "MIS"]].sum()
            st.plotly_chart(daily_trend_fig(trend), use_container_width=True, key="chart_day_trend", config=PLOTLY_CONFIG)

    site_summary = compute_site_summary(df, threshold)
    left, right = st.columns([1.65, 1], gap="medium")
    with left:
        with st.container(border=True, key="card_2"):
            top = site_summary.head(10)
            card_header("Sites with the largest difference",
                        f"Top {len(top)} of {len(site_summary)} sites by MIS vs GPS difference — dotted lines mark the ±{threshold}% limit")
            st.plotly_chart(site_deviation_fig(top, threshold, max(300, len(top) * 34 + 60)),
                            use_container_width=True, key="chart_overview_sites", config=PLOTLY_CONFIG)
    with right:
        with st.container(border=True, key="card_3"):
            fig_status, legend = fleet_status_fig(df)
            card_header("Fleet status", "Vehicles by recommended action", legend=legend)
            st.plotly_chart(fig_status, use_container_width=True, key="chart_fleet_status", config=PLOTLY_CONFIG)

    all_months = sorted(load_history().values(), key=lambda e: e["key"])
    if len(all_months) >= 2:
        with st.container(border=True, key="card_4"):
            card_header("Month-on-month totals", "Total GPS vs MIS km for every saved month",
                        legend=[("GPS", COLOR_GPS), ("MIS", COLOR_MIS)])
            st.plotly_chart(monthly_totals_fig(all_months, 280), use_container_width=True,
                            key="chart_overview_months", config=PLOTLY_CONFIG)

    if not has_daily:
        st.caption("The day-wise chart isn't available for this month — it was saved before daily figures were stored. Re-upload the file to add it.")


def compute_site_summary(df, threshold):
    site_summary = df.groupby("Site", as_index=False).agg(
        Vehicles=("Vehicle", "count"), Total_GPS=("Total GPS", "sum"), Total_MIS=("Total MIS", "sum"),
        Flagged=("Action", lambda s: (s != "No action needed").sum()),
    )
    site_summary["Diff %"] = ((site_summary["Total_MIS"] - site_summary["Total_GPS"]) / site_summary["Total_GPS"].replace(0, pd.NA)).fillna(0)
    site_summary = site_summary.sort_values("Diff %", key=abs, ascending=False)

    def site_severity(diff_pct, threshold_pct):
        t = threshold_pct / 100
        if abs(diff_pct) > t:
            return "critical"
        if abs(diff_pct) > t * 0.5:
            return "watch"
        return "ok"

    site_summary["Severity"] = site_summary["Diff %"].apply(lambda d: site_severity(d, threshold))
    return site_summary


def page_sites():
    df = st.session_state.df
    threshold = st.session_state.threshold

    render_brand_header("Sites", "Kilometre volume and GPS vs MIS difference for every site.")

    site_summary = compute_site_summary(df, threshold)
    n_sites = len(site_summary)
    n_crit = int((site_summary["Severity"] == "critical").sum())
    n_watch = int((site_summary["Severity"] == "watch").sum())

    c1, c2, c3, c4 = st.columns(4)
    kpi_card(c1, "building", COLOR_VEHICLES, "Sites", f"{n_sites}", vs_text=f"{len(df):,} vehicles in total")
    kpi_card(c2, "flag", COLOR_CRIT, "Outside limit", f"{n_crit}",
             delta_text=f"{n_crit/n_sites*100:.0f}% of sites" if n_sites else None, delta_tone="bad" if n_crit else "good",
             meter=(n_crit / n_sites) if n_sites else 0)
    kpi_card(c3, "sliders", COLOR_WARN, "On watch", f"{n_watch}", vs_text=f"between ±{threshold/2:g}% and ±{threshold}%",
             meter=(n_watch / n_sites) if n_sites else 0)
    worst = site_summary.iloc[0] if n_sites else None
    kpi_card(c4, "diff", COLOR_DIFF, "Largest difference",
             f"{worst['Diff %']*100:+.1f}" if worst is not None else "—", unit="%" if worst is not None else None,
             vs_text=str(worst["Site"]) if worst is not None else None)

    st.write("")
    with st.container(border=True, key="card_5"):
        card_header("Difference by site", f"MIS vs GPS km difference — dotted lines mark the ±{threshold}% flag limit",
                    legend=[(STATUS_LABEL[s], STATUS_FILL[s]) for s in ("critical", "watch", "ok")])
        st.plotly_chart(site_deviation_fig(site_summary, threshold, max(320, n_sites * 30 + 60)),
                        use_container_width=True, key="chart_sites_diff", config=PLOTLY_CONFIG)

    vol_sorted = site_summary.sort_values("Total_GPS", ascending=True)
    fig_vol = go.Figure()
    for col, name, color in (("Total_GPS", "GPS", COLOR_GPS), ("Total_MIS", "MIS", COLOR_MIS)):
        fig_vol.add_trace(go.Bar(
            y=vol_sorted["Site"], x=vol_sorted[col], name=name, orientation="h",
            marker=dict(color=color, cornerradius=4),
            hovertemplate=f"<b>%{{y}}</b><br>{name}: %{{x:,.0f}} km<extra></extra>",
        ))
    style_fig(fig_vol, max(320, n_sites * 34), x_title="Km")
    fig_vol.update_xaxes(showgrid=True, gridcolor=CHART_GRID, tickformat=",.0f")
    fig_vol.update_yaxes(showgrid=False, automargin=True, tickfont=dict(size=12.5, color=COLOR_TEXT))
    fig_vol.update_layout(barmode="group", bargap=0.3, bargroupgap=0.1)

    with st.container(border=True, key="card_6"):
        card_header("Kilometre volume by site", "Total GPS vs MIS km, largest sites at the top — scroll inside the box to see all sites",
                    legend=[("GPS", COLOR_GPS), ("MIS", COLOR_MIS)])
        with st.container(height=440, border=False):
            st.plotly_chart(fig_vol, use_container_width=True, key="chart_sites_vol", config=PLOTLY_CONFIG)

    def style_site_row(row):
        color = ACTION_COLORS.get(row["Severity"], COLOR_TEXT)
        return [f"color: {color}; font-weight: 600;" if col == "Diff %" else "" for col in row.index]

    site_styled = (
        site_summary.style
        .apply(style_site_row, axis=1)
        .format({"Total_GPS": "{:,.0f}", "Total_MIS": "{:,.0f}", "Diff %": "{:+.1%}"})
        .hide(axis="columns", subset=["Severity"])
    )
    st.markdown("#### Site table")
    st.dataframe(site_styled, use_container_width=True, hide_index=True, column_config={"Severity": None})


def page_vehicles():
    df = st.session_state.df

    render_brand_header("Vehicles", "Recommended corrective action for every vehicle — click a tile to filter the table.")

    st.markdown("### Corrective actions")
    order = ["Check GPS device", "File missing MIS log", "No data either side", "Audit MIS entries",
             "Verify unrecorded trips", "Keep an eye on it", "No action needed"]
    action_counts = df.groupby(["Action", "Severity"]).size().reset_index(name="Count")
    action_counts["order"] = action_counts["Action"].apply(lambda a: order.index(a) if a in order else 99)
    action_counts = action_counts.sort_values("order")

    if "action_filter" not in st.session_state:
        st.session_state.action_filter = "All"

    SEVERITY_EMOJI = {"critical": "🔴", "warn": "🟠", "watch": "🟡", "ok": "🟢"}

    with st.container(key="corrective_actions"):
        cols = st.columns(len(action_counts)) if len(action_counts) else [st]
        for idx, (col, (_, row)) in enumerate(zip(cols, action_counts.iterrows())):
            is_active = st.session_state.action_filter == row["Action"]
            emoji = SEVERITY_EMOJI.get(row["Severity"], "⚪")
            label = f"{emoji} {row['Count']}\n{row['Action']}"
            if is_active:
                label = f"✓ {emoji} {row['Count']}\n{row['Action']}"
            with col:
                with st.container(key=f"actsev_{row['Severity']}_{idx}"):
                    if st.button(label, key=f"actbtn_{row['Action']}", use_container_width=True):
                        st.session_state.action_filter = "All" if is_active else row["Action"]
                        st.rerun()

    st.markdown("### Vehicles")
    if st.session_state.action_filter != "All":
        fc1, fc2 = st.columns([5, 1])
        fc1.info(f"Filtered to action: **{st.session_state.action_filter}**")
        if fc2.button("✕ Clear filter", use_container_width=True):
            st.session_state.action_filter = "All"
            st.rerun()

    col_a, col_b, col_c, col_d = st.columns([1.3, 2, 1, 1.3])
    site_options = ["All sites"] + sorted(df["Site"].unique().tolist())
    selected_site = col_a.selectbox("Site", site_options)
    search = col_b.text_input("Search vehicle no. or site")
    only_flagged = col_c.checkbox("Flagged only")
    action_filter = col_d.selectbox("Action", ["All"] + order, key="action_filter")

    view = df.copy()
    if selected_site != "All sites":
        view = view[view["Site"] == selected_site]
    if only_flagged:
        view = view[view["Action"] != "No action needed"]
    if action_filter != "All":
        view = view[view["Action"] == action_filter]
    if search.strip():
        q = search.strip().lower()
        view = view[view["Vehicle"].str.lower().str.contains(q) | view["Site"].str.lower().str.contains(q)]

    view = view.sort_values("Diff %", key=abs, ascending=False)
    st.caption(f"{len(view)} vehicles")

    display_cols = ["Vehicle", "Site", "Total GPS", "Total MIS", "Diff %", "Action", "Source", "Severity"]
    display_df = view[display_cols]

    ACTION_BG = {
        "critical": hex_to_rgba(COLOR_CRIT, 0.14),
        "warn": hex_to_rgba(COLOR_WARN, 0.14),
        "watch": hex_to_rgba(COLOR_WARN, 0.08),
        "ok": hex_to_rgba(COLOR_OK, 0.10),
    }

    def style_row(row):
        color = ACTION_COLORS.get(row["Severity"], COLOR_TEXT)
        bg = ACTION_BG.get(row["Severity"], "")
        styles = []
        for col in row.index:
            if col == "Diff %":
                styles.append(f"color: {color}; font-weight: 600;")
            elif col == "Action":
                styles.append(f"color: {color}; background-color: {bg}; font-weight: 500;")
            else:
                styles.append("")
        return styles

    styled = (
        display_df.style
        .apply(style_row, axis=1)
        .format({"Total GPS": "{:,.1f}", "Total MIS": "{:,.1f}", "Diff %": "{:+.1%}"})
        .hide(axis="columns", subset=["Severity"])
    )

    st.dataframe(styled, use_container_width=True, hide_index=True, height=460, column_config={"Severity": None})


def page_drilldown():
    df = st.session_state.df
    daily_df = st.session_state.daily_df
    has_daily = st.session_state.has_daily
    threshold = st.session_state.threshold

    render_brand_header("Vehicle drill-down", "Month totals, recommended action and the day-by-day record for one vehicle.")
    veh_pick = st.selectbox("Pick a vehicle", df["Vehicle"].tolist())
    if not veh_pick:
        return
    row = df[df["Vehicle"] == veh_pick].iloc[0]
    color = ACTION_COLORS[row["Severity"]]

    c1, c2, c3, c4 = st.columns(4)
    kpi_card(c1, "gps", COLOR_GPS, "GPS km", f"{row['Total GPS']:,.1f}", vs_text=f"Site: {row['Site']}")
    kpi_card(c2, "mis", COLOR_MIS, "MIS km", f"{row['Total MIS']:,.1f}", vs_text=f"Source: {row['Source']}")
    over = abs(row["Diff %"]) * 100 > threshold
    kpi_card(c3, "diff", COLOR_DIFF, "Difference", f"{row['Diff %']*100:+.1f}", unit="%",
             delta_text=f"{row['Diff']:+,.1f} km", delta_tone="bad" if over else "good",
             vs_text=f"{'Outside' if over else 'Within'} ±{threshold}% limit")
    kpi_card(c4, "flag", color, "Status", STATUS_LABEL.get(row["Severity"], "—"), vs_text=row["Action"])

    st.markdown(
        f"<div class='action-card' style='border-left:4px solid {color}; margin-top:14px;'>"
        f"<b style='color:{color};'>{row['Action']}</b><br>"
        f"<span style='color:{COLOR_MUTED}; font-size:13.5px;'>{row['ActionDetail']}</span></div>",
        unsafe_allow_html=True,
    )
    vd = daily_df[daily_df["Vehicle"] == veh_pick]
    if has_daily and len(vd):
        with st.container(border=True, key="card_7"):
            card_header(f"Daily kilometres — {veh_pick}", "Km recorded each day by GPS and logged in MIS",
                        legend=[("GPS", COLOR_GPS), ("MIS", COLOR_MIS)])
            st.plotly_chart(daily_trend_fig(vd, value_fmt=",.1f"), use_container_width=True,
                            key="chart_vehicle_drilldown", config=PLOTLY_CONFIG)
    else:
        st.caption("The day-wise chart isn't available for this month — it was saved before daily figures were stored.")


def page_history():
    render_brand_header("Monthly history", "Every uploaded month is saved automatically, so you can track the trend.")

    hist = load_history()  # reload in case this run just saved a new entry
    if not hist:
        st.info("No saved months yet — upload a file to start building history.")
        return

    sorted_entries = sorted(hist.values(), key=lambda x: x["key"])
    hist_df = pd.DataFrame([{
        "Month": e["label"], "Total GPS": e["total_gps"], "Total MIS": e["total_mis"],
        "Diff %": ((e["total_mis"] - e["total_gps"]) / e["total_gps"]) if e["total_gps"] else 0,
        "Vehicles": e["vehicles"], "Flagged": e.get("flagged", "-"),
    } for e in sorted_entries])

    threshold = st.session_state.threshold
    left, right = st.columns([1.5, 1], gap="medium")
    with left:
        with st.container(border=True, key="card_8"):
            card_header("Total kilometres by month", "GPS vs MIS km for each saved month",
                        legend=[("GPS", COLOR_GPS), ("MIS", COLOR_MIS)])
            st.plotly_chart(monthly_totals_fig(sorted_entries, 300), use_container_width=True,
                            key="chart_monthly_history", config=PLOTLY_CONFIG)
    with right:
        with st.container(border=True, key="card_9"):
            card_header("Overall difference by month", f"MIS vs GPS — dotted lines mark the ±{threshold}% limit",
                        legend=[(STATUS_LABEL[s], STATUS_FILL[s]) for s in ("critical", "watch", "ok")])
            st.plotly_chart(monthly_diff_fig(sorted_entries, threshold, 300), use_container_width=True,
                            key="chart_monthly_diff", config=PLOTLY_CONFIG)

    st.dataframe(
        hist_df.style.format({"Total GPS": "{:,.0f}", "Total MIS": "{:,.0f}", "Diff %": "{:+.1%}"}),
        use_container_width=True, hide_index=True,
    )

    with st.expander("Remove a saved month"):
        del_choice = st.selectbox("Month to remove", [e["label"] for e in sorted_entries])
        if st.button("Delete this month from history"):
            key_to_delete = next(e["key"] for e in sorted_entries if e["label"] == del_choice)
            del hist[key_to_delete]
            save_history(hist)
            st.success(f"Removed {del_choice}. Refresh the page to see the updated list.")


# ---------------------------------------------------------------------------
# Navigation
# ---------------------------------------------------------------------------
pg = st.navigation([
    st.Page(page_overview, title="Overview", icon="📊", url_path="overview", default=True),
    st.Page(page_sites, title="Sites", icon="🏢", url_path="sites"),
    st.Page(page_vehicles, title="Vehicles", icon="🚚", url_path="vehicles"),
    st.Page(page_drilldown, title="Vehicle Drill-down", icon="🔍", url_path="drilldown"),
    st.Page(page_history, title="Monthly History", icon="📅", url_path="history"),
])
with st.sidebar:
    st.markdown("#### Report")
    report_download_button("dl_sidebar")

pg.run()
render_footer()
