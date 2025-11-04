# streamlit_app.py

import streamlit as st
import pandas as pd
import math
import io
from datetime import datetime, timedelta
import plotly.express as px

# ========================================
# PAGE CONFIGURATION
# ========================================
st.set_page_config(
    page_title="Cyclo Production Planning AI",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ========================================
# STYLES
# ========================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    .hero-header {
        background: linear-gradient(135deg, #63913A 0%, #7AB850 100%);
        padding: 3rem 2rem;
        border-radius: 20px;
        text-align: center;
        color: white;
        margin-bottom: 3rem;
        box-shadow: 0 10px 30px rgba(99, 145, 58, 0.3);
    }
    .upload-card { padding: 2rem; border-radius: 12px; border: 2px solid #e8f5e9; box-shadow: 0 4px 6px rgba(0,0,0,0.07); text-align:center; }
    .stButton > button {
        background: linear-gradient(135deg, #63913A 0%, #7AB850 100%);
        color: white; border: none; border-radius: 12px; padding: 1rem 3rem;
        font-weight: 600; font-size: 1.1rem; transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(99, 145, 58, 0.3);
    }
    .stButton > button:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(99, 145, 58, 0.4); }
    .metric-card {  padding: 1.2rem; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); border-left: 5px solid #63913A; }
    .metric-label { font-size: 0.85rem; color:#666; text-transform: uppercase; margin-bottom:0.4rem; font-weight:600; }
    .metric-value { font-size:1.6rem; font-weight:700; color:#63913A; }
    .section-header { font-size:1.4rem; font-weight:700; color:#333; margin:1.2rem 0 0.8rem 0; padding-bottom:0.4rem; border-bottom:3px solid #63913A; }
    .dataframe thead tr th { background: #63913A !important; color: white !important; font-weight: 600 !important; padding: 0.6rem !important; }
</style>
""", unsafe_allow_html=True)

# ========================================
# CONSTANTS & CONFIGURATIONS
# ========================================

# --- Aliases for differing column labels between testing vs. actual files ---
COLUMN_ALIASES = {
    "PI NO": ["PI NO", "PI No", "PI Number", "PI#", "Invoice No"],
    "Yarn Count": ["Yarn Count", "Count", "Count (Ne)", "Ne Count", "Count Ne", "Count Ne/"],
    "Composition": ["Composition", "Blend", "Material Composition"],
    "Yarn Type": ["Yarn Type", "Type", "YarnType"],
    "Color Code": ["Color Code", "Colour Code", "ColorCode", "Shade Code"],
    "ColorFamilyName": ["ColorFamilyName", "Color Family", "Colour Family", "Family Color"],
    "Quantity": ["Quantity", "Qty", "Quantity (KG)", "Quantity (kg)", "QTY (KG)", "QTY (kg)", "Order Qty (KG)"],
    "Due Date": ["Due Date", "Delivery Date", "Customer requested Delivery Date"],
    # Optional helper columns:
    "Color": ["Color", "Shade", "Colour"],
    "Invoice Date": ["Invoice Date", "PI Date"],
    "Customer": ["Customer", "Buyer"]
}

REQUIRED_STD_COLS = ["PI NO", "Yarn Count", "Composition", "Yarn Type", "Quantity"]

# Composition normalization map
BLEND_MAPPING = {
    "70% CYCLO® Recycled Cotton 30% Recycled Polyester": "70/30 CYL Cot/poly",
    "70% CYCLO® Recycled Cotton 30% Polyester": "70/30 CYL Cot/poly",
    "30% Polyester 70% CYCLO® Recycled Cotton": "70/30 CYL Cot/poly",
    "70% CYCLO® Recycled Cotton 30% Recycled Polyester 0.001% Tracer Fibers": "70/30 CYL Cot/poly",
    "80% CYCLO® Recycled Cotton 20% Recycled Polyester": "80/20 CYL Cot/poly",
    "90% CYCLO® Recycled Cotton 10% Recycled Polyester": "90/10 CYL Cot/poly",
    "90% CYCLO® Recycled Cotton 10% Polyesterr": "90/10 CYL Cot/poly",
    "60% CYCLO® Recycled Cotton 40% Recycled Polyester": "60/40 CYL Cot/poly",
    "60% CYCLO® Recycled Cotton 40% Polyester": "60/40 CYL Cot/poly",
    "50% CYCLO® Recycled Cotton 50% Recycled Polyester": "50/50 CYL Cot/poly",
    "80% CYCLO® Recycled Cotton 20% Recycled Polyester 0.001% Tracer Fibers": "80/20 CYL Cot/poly Fiber tracer 0.001%",
    "70% CYCLO® Recycled Cotton 30% Recycled Polyester 0.001% Tracer Fibers": "70/30 CYL Cot/poly Fiber tracer 0.001%",
    "50% CYCLO® Recycled Cotton 30% Recycled Polyester 20% Nylon": "50/30/20 CYL Cot/poly/nylon",
    "50% CYCLO® Recycled Cotton 50% Recycled Polyester 0.001% Tracer Fibers": "50/50 CYL Cot/poly Fiber tracer 0.001%",
    "50% CYCLO® Recycled Cotton  50% ECOVERO™ Viscose":"50/50 ECOVERO™ Viscose/Cyl Cot",
    "50% ECOVERO™ Viscose 50% CYCLO® Recycled Cotton  ":"50/50 ECOVERO™ Viscose/Cyl Cot",
    "50% CYCLO® Recycled Cotton   50% Liva Reviva™ Viscose":"50/50 Liva Reviva™ Viscose/Cyl Cot",
    "50% Liva Reviva™ Viscose 50% CYCLO® Recycled Cotton":"50/50 Liva Reviva™ Viscose/Cyl Cot",
    "50% CYCLO® Recycled Cotton 50% Lyocell":"50/50 Lyocell/Cyl Cot",
    "50% Lyocell 50% CYCLO® Recycled Cotton ":"50/50 Lyocell/Cyl Cot",
    "50% CYCLO® Recycled Cotton   50% Organic Cotton":"50/50 Org/Cyl Cot",
    "50% Organic Cotton 50% CYCLO® Recycled Cotton":"50/50 Org/Cyl Cot",
    "50% CYCLO® Recycled Cotton   50% Virgin Cotton":"50/50 CYL Cot/Virgin Cot",
    "50% Virgin Cotton 50% CYCLO® Recycled Cotton":"50/50 CYL Cot/Virgin Cot",
    "55% CYCLO® Recycled Cotton   30% Virgin Cotton 15% Polyester":"55/30/15 CYL Cot/Virgin Cot/poly",
    "60% CYCLO® Recycled Cotton   20% Viscose 20% Nylon":"60/20/20 CYL Cot/Viscose/Nylon",
    "60% CYCLO® Recycled Cotton   40% Recycled Polyester":"60/40 CYL Cot/poly",
    "70% CYCLO® Recycled Cotton   30% Acrylic":"70/30 CYL Cot/Acrylic",
    "50% CYCLO® Recycled Cotton 50% Bamboo":"50/50 Bamboo/Cyl Cot",
    "50% Bamboo 50% CYCLO® Recycled Cotton":"50/50 Bamboo/Cyl Cot",
    "70% CYCLO® Recycled Cotton   30% Bamboo Viscose":"30/70 Bamboo Viscose/Cyl Cot",
    "30% Bamboo Viscose 70% CYCLO® Recycled Cotton":"30/70 Bamboo Viscose/Cyl Cot",
    "70% CYCLO® Recycled Cotton   30% Liva Reviva™ Viscose":"30/70 Liva Reviva™ Viscose/Cyl Cot",
    "30% Liva Reviva™ Viscose 70% CYCLO® Recycled Cotton":"30/70 Liva Reviva™ Viscose/Cyl Cot",
    "90% CYCLO® Recycled Cotton   10% Tencel™":"90/10 CYL Cot/Tencel™",
    "10% Tencel™ 90% CYCLO® Recycled Cotton":"90/10 CYL Cot/Tencel™",
    "60% CYCLO® Recycled Cotton 40% Polyester":"60/40 CYL Cot/poly",
    "40% Polyester 60% CYCLO® Recycled Cotton ":"60/40 CYL Cot/poly",
    "70% Organic Cotton 30% CYCLO® Recycled Cotton":"70/30 Org/Cyl Cot",
    "30% CYCLO® Recycled Cotton 70% Organic Cotton":"70/30 Org/Cyl Cot",
    "52% CYCLO® Recycled Cotton 27% Polyester 21% Nylon":"52/27/23 CYL Cot/poly/nylon",
    "70% CYCLO® Recycled Cotton 30% Lyocell":"30/70 Lyocell/Cyl Cot",
    "30% Lyocell 70% CYCLO® Recycled Cotton ":"30/70 Lyocell/Cyl Cot",
    "70% CYCLO® Recycled Cotton 20% Linen 10% Viscose":"70/20/10 CYL Cot/Linen/Viscose",
    "70% CYCLO® Recycled Cotton 30% Recycled Polyester (BPA free)":"70/30 CYL Cot/poly",
    "30% Recycled Polyester 70% CYCLO® Recycled Cotton (BPA free)":"70/30 CYL Cot/poly",    
    "45% CYCLO® Recycled Cotton 40% Recycled Polyester 15% Post Consumer Recycled Cotton Polyester":"45/40/15 CYL Cot/ploy/PCW Cot/poly" ,
    "50% Organic Cotton 35% CYCLO® Recycled Cotton 15% Post Consumer Recycled Cotton":"45/40/15 CYL Cot/ploy/PCW Cot/poly",
    "50% BCI Cotton 50% CYCLO® Recycled Cotton":"50/50 BCI/Cyl Cot",
    "50% CYCLO® Recycled Cotton 50% BCI Cotton":"50/50 BCI/Cyl Cot",
    "65% BCI Cotton 20% CYCLO® Recycled Cotton 15% Ecovero Viscose Mélange":"65/20/15 BCI Cot/CYL cot/Ecovero  Viscose Mélange",
    "80% BCI Cotton 20% CYCLO® Recycled Cotton":"80/20 BCI/Cyl Cot",
    "20% CYCLO® Recycled Cotton 80% BCI Cotton ":"80/20 BCI/Cyl Cot",
    "70% BCI Cotton 30% CYCLO® Recycled Cotton":"70/30 BCI/Cyl Cot",
    "30% CYCLO® Recycled Cotton 70% BCI Cotton ":"70/30 BCI/Cyl Cot",
    "50% CYCLO® Recycled Cotton 50% Acrylic":"50/50 CYL Cot/Acrylic",
    "50% Acrylic 50% CYCLO® Recycled Cotton":"50/50 CYL Cot/Acrylic",
    "50% CYCLO® Recycled Cotton   50% Recycled Polyester":"50/50 CYL Cot/poly",
    "50% Recycled Polyester 50% CYCLO® Recycled Cotton":"50/50 CYL Cot/poly",
    "50% CYCLO® Recycled Cotton 50% Viscose":"50/50 Bamboo Viscose/Cyl Cot",
    "50% Viscose 50% CYCLO® Recycled Cotton":"50/50 Bamboo Viscose/Cyl Cot",
    "55% CYCLO® Recycled Cotton   45% Recycled Polyester":"55/45 CYL Cot/poly",
    "45% Recycled Polyester 55% CYCLO® Recycled Cotton ":"55/45 CYL Cot/poly",
    "60% CYCLO® Recycled Cotton  40% Nylon":"60/40 CYL Cot/nylon",
    " 40% Nylon 60% CYCLO® Recycled Cotton  ":"60/40 CYL Cot/nylon",
    "60% CYCLO® Recycled Cotton   40% Virgin Cotton":"60/40 CYL Cot/Virgin Cot",
    "40% Virgin Cotton 60% CYCLO® Recycled Cotton":"60/40 CYL Cot/Virgin Cot",
    "65% CYCLO® Recycled Cotton 35% Organic Cotton Tracer Fibers":"35/65 BCI/Cyl Cot",
    "65% CYCLO® Recycled Cotton 35% Polyester":" 65/35 CYL Cot/poly",
    "35% Polyester 65% CYCLO® Recycled Cotton":" 65/35 CYL Cot/poly",
    "70% CYCLO® Recycled Cotton   30% Virgin Cotton":"30/70 Virgin Cot/Cyl Cot",
    "30% Virgin Cotton 70% CYCLO® Recycled Cotton":"30/70 Virgin Cot/Cyl Cot",
    "70% CYCLO® Recycled Cotton   30% Viscose":"30/70  Viscose/Cyl Cot",
    "30% Viscose 70% CYCLO® Recycled Cotton":"30/70  Viscose/Cyl Cot",
    "97% CYCLO® Recycled Cotton   3% Recycled Polyester":"97/3 CYL Cot/poly",
    "3% Recycled Polyester 97% CYCLO® Recycled Cotton":"97/3 CYL Cot/poly",
    "80% Organic Cotton 20% Recycled Polyester":"80/20 Org Cot/poly",
    "20% Recycled Polyester 80% Organic Cotton":"80/20 Org Cot/poly",
    "50% CYCLO® Recycled Cotton 30% BCI Cotton 20% Recycled Polyester":"50/20/30 CYL Cot/BCI Cot/poly",
    "30% BCI Cotton 20% Recycled Polyester 50% CYCLO® Recycled Cotton":"50/20/30 CYL Cot/BCI Cot/poly",
    "65% CYCLO® Recycled Cotton 35% Recycled Polyester":"65/35 CYL Cot/poly",
    "35% Recycled Polyester 65% CYCLO® Recycled Cotton":"65/35 CYL Cot/poly",
    "65% Recycled Polyester 35% BCI Cotton":"65/35 poly/BCI",
    "35% BCI Cotton 65% Recycled Polyester":"65/35 poly/BCI",
    "60% Organic Cotton 40% Polyester":"60/40 Org/poly",
    "40% Polyester 60% Organic Cotton":"60/40 Org/poly",
    "73% CYCLO® Recycled Cotton 25% Recycled Polyester 2% Viscose":"73/25/2 CYL Cot/poly/Viscose",
    "65% Polyester 35% Virgin Cotton":"65/35 poly/Virgin Cot",
    "35% Virgin Cotton 65% Polyester":"65/35 poly/Virgin Cot",
    "75% CYCLO® Recycled Cotton 25% Polyester":"75/25 CYL Cot/poly",
    "25% Polyester 75% CYCLO® Recycled Cotton":"75/25 CYL Cot/poly",
    "60% Organic Cotton 40% Recycled Polyester":"60/40 Org/poly",
    "40% Recycled Polyester 60% Organic Cotton":"60/40 Org/poly",
    "60% CYCLO® Recycled Cotton 30% Recycled Polyester 10% Viscose":"60/30/10 CYL Cot/poly/Viscose",
    "60% CYCLO® Recycled Cotton 40% Acrylic":" 60/40 CYL Cot/Acrylic",
    "40% Acrylic 60% CYCLO® Recycled Cotton":" 60/40 CYL Cot/Acrylic",
    "55% CYCLO® Recycled Cotton 25% Polyester 20% Nylon":"55/25/20 CYL Cot/poly/nylon",
    "25% Polyester 20% Nylon 55% CYCLO® Recycled Cotton":"55/25/20 CYL Cot/poly/nylon",
    "70% CYCLO® Recycled Cotton 30% Viscose":"30/70  Viscose/Cyl Cot",
    "30% Viscose 70% CYCLO® Recycled Cotton":"30/70  Viscose/Cyl Cot",
    "75% CYCLO® Recycled Cotton 25% Recycled Polyester":"75/25 CYL Cot/poly",
    "25% Recycled Polyester 75% CYCLO® Recycled Cotton":"75/25 CYL Cot/poly",
    "65% Recycled Polyester 35% Organic Cotton":"65/35 poly/Org",
    "35% Organic Cotton 65% Recycled Polyester":"65/35 poly/Org",
    "60% BCI Cotton 40% Recycled Polyester":"60/40 BCI/poly",
    "40% Recycled Polyester 60% BCI Cotton":"60/40 BCI/poly",
    "70% Virgin Cotton 30% CYCLO® Recycled Cotton":" 70/30 Virgin/Cyl Cot",
    "30% CYCLO® Recycled Cotton 70% Virgin Cotton":" 70/30 Virgin/Cyl Cot",
    "60% Virgin Cotton 40% CYCLO® Recycled Cotton":"60/40 Virgin/Cyl Cot",
    "40% CYCLO® Recycled Cotton 60% Virgin Cotton":"60/40 Virgin/Cyl Cot",
    "75% BCI Cotton 20% CYCLO® Recycled Cotton 5% Ecovero Viscose Mélange":"75/20/5 BCI Cot/CYL cot/Ecovero  Viscose Mélange",
    "70% CYCLO® Recycled Cotton 30% Nylon":" 70/20 CYL Cot/nylon",
    "30% Nylon 70% CYCLO® Recycled Cotton":" 70/20 CYL Cot/nylon",
    "75% Organic Cotton 25% CYCLO® Recycled Cotton":"75/25 BCI/Cyl Cot",
    "25% CYCLO® Recycled Cotton 75% Organic Cotton":"75/25 BCI/Cyl Cot",
    "50% Recycled Polyester 35% CYCLO® Recycled Cotton 15% Post Consumer Recycled Cotton":"35/50/15 CYL Cot/ploy/PCW Cot",
    "65% Recycled Polyester 35% CYCLO® Recycled Cotton":"35/65 CYL Cot/poly",
    "35% CYCLO® Recycled Cotton 65% Recycled Polyester":"35/65 CYL Cot/poly",
    "50% Recycled Polyester 35% CYCLO® Recycled Cotton 15% Post Consumer (65% Cotton 35% Polyester)":"35/50/15 CYL Cot/ploy/PCW Cot/poly",  
    "50% CYCLO® Recycled Cotton 50% Organic Cotton":"50/50 Org/Cyl Cot",
    "50% Organic Cotton 50% CYCLO® Recycled Cotton":"50/50 Org/Cyl Cot",
    "60% CYCLO® Recycled Cotton 20% Viscose 20% Nylon":"60/20/20 CYL Cot/Viscose/Nylon",
    "100% CYCLO® Recycled Cotton":"100 CYL Cot",
    "50% CYCLO® Recycled Cotton 25% Recycled Polyester 25% Viscose":"50/25/25 CYL Cot/poly/Viscose",
}

NEAREST_FAMILIES = {
    "RED": ["Maroon", "Rust Melange", "Orange"],
    "Green": ["Midnight Olive", "Pearl Teal"],
    "Blue": ["Denim", "Midnight Blue", "Pearl Teal"],
    "Beige": ["Stone", "Cream", "Natural"],
    "White": ["Cream", "Grey"],
    "Yellow": ["Dijon", "Cream"],
    "Stone": ["Beige", "Natural", "Grey", "Anthracite" ],
    "Anthracite" : ["Stone", "Grey"],
    "Midnight Olive": ["Green", "Brown"],
    "Golden Mocha": ["Brown", "Chocolate"],
    "Charcoal": ["Grey", "Black"],
    "Midnight Blue": ["Blue", "Denim"],
    "Pearl Teal": ["Aqua", "Turquoise", "Green"],
    "Maroon": ["RED", "Rust Melange", "Brown"],
    "Brown": ["Chocolate", "Golden Mocha", "Maroon"],
    "Rust Melange": ["Maroon", "RED", "Brown"],
    "Rose": ["Pink", "Magenta"],
    "Pink": ["Rose", "Magenta", "RED"],
    "Grey": ["Charcoal", "White", "Stone"],
    "Purple": ["Magenta", "Pink"],
    "Chocolate": ["Brown", "Golden Mocha"],
    "Aqua": ["Turquoise", "Pearl Teal"],
    "Black": ["Charcoal", "Grey"],
    "Cream": ["Beige", "White", "Natural"],
    "Denim": ["Blue", "Midnight Blue"],
    "Dijon": ["Yellow", "Brown"],
    "Natural": ["Beige", "Cream", "Stone"],
    "Orange": ["RED", "Rust Melange"],
    "Turquoise": ["Aqua", "Pearl Teal"],
    "Magenta": ["Pink", "Rose", "Purple"],
}

# ---------- Real five-line configuration ----------
LINE_CONFIG = {
  "Line 1": {"machines": 4, "spindles_per_machine": 460, "daily_capacity_kg": 5000},
  "Line 2": {"machines": 3, "spindles_per_machine": 460, "daily_capacity_kg": 5000},
  "Line 3": {"machines": 3, "spindles_per_machine": 460, "daily_capacity_kg": 5000},
  "Line 4": {"machines": 2, "spindles_per_machine": 240, "daily_capacity_kg": 2000},
  "Line 5": {"machines": 2, "spindles_per_machine": 240, "daily_capacity_kg": 2000},
}
LINES_MAIN  = ["Line 1", "Line 2", "Line 3"]
LINES_SMALL = ["Line 4", "Line 5"]
LINES = list(LINE_CONFIG.keys())

SHIFTS = [("A", 0, 480), ("B", 480, 960), ("C", 960, 1440)]
SHIFT_DURATION_MIN = 480
HORIZON_DAYS = 60

# Per-line shift capacity (kg per shift)
PER_SHIFT_CAPACITY = {ln: LINE_CONFIG[ln]["daily_capacity_kg"] / len(SHIFTS) for ln in LINES}

# MACHINE_FILE_PATH = "./reports/machine.xlsx"
MACHINE_FILE_PATH = "./reports/updated_machine_data.xlsx"

# ========================================
# HELPER FUNCTIONS
# ========================================
def round_up(val: float) -> float:
    return math.ceil(val * 100) / 100

def _alias_to_std(name: str) -> str:
    name = (name or "").strip()
    for std, alts in COLUMN_ALIASES.items():
        if name in alts:
            return std
    return name

def _detect_header_row(df_noheader: pd.DataFrame) -> int:
    target_aliases = {a for alts in COLUMN_ALIASES.values() for a in alts}
    best_row, best_score = 0, 0
    for r in range(min(60, len(df_noheader))):
        vals = [str(x).strip() for x in df_noheader.iloc[r].tolist()]
        score = sum(1 for v in vals if v in target_aliases)
        if score > best_score:
            best_row, best_score = r, score
    return best_row

def _standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    new_cols = []
    for c in df.columns:
        new_cols.append(_alias_to_std(str(c)))
    df.columns = new_cols
    return df

def normalize_count(count_value):
    s = str(count_value).strip()
    if not s or s.lower() == "nan":
        return None
    if "/" in s:
        part = s.split("/")[0]
    else:
        part = s.split()[0]
    try:
        return int(float(part))
    except:
        return None

def _clean_text(s):
    return " ".join(str(s).split())

def normalize_blend(blend_raw: str) -> str:
    clean_blend = _clean_text(blend_raw)
    mapped = BLEND_MAPPING.get(clean_blend)
    if mapped:
        return mapped
    try2 = clean_blend.replace("®", "").replace("™", "")
    try2 = " ".join(try2.split())
    return BLEND_MAPPING.get(try2, None)

def ensure_date(dt):
    if pd.isna(dt):
        return None
    if isinstance(dt, str):
        return pd.to_datetime(dt, errors='coerce').date()
    if isinstance(dt, pd.Timestamp):
        return dt.date()
    if isinstance(dt, datetime):
        return dt.date()
    return dt

def calculate_hours(order, machines):
    """
    Returns (estimated_hours, error) based on LINE_CONFIG & machine table.
    For 500–2000 kg we choose the faster of Lines 4–5; otherwise faster of Lines 1–3.
    """
    try:
        count = normalize_count(order["Yarn Count"])
        blend_raw = str(order["Composition"]).strip()
        yarn_type = str(order["Yarn Type"]).strip()
        qty_required = float(order["Quantity"])

        blend = normalize_blend(blend_raw)
        if not blend:
            return None, f"Blend not mapped: {blend_raw}"

        matched = machines[
            (machines["Counts"] == count) &
            (machines["Blends"] == blend) &
            (machines["Yarn Type"] == yarn_type)
        ]

        if matched.empty:
            return None, "No machine data"

        best_row = matched.loc[matched["twist factor"].idxmax()]
        twist_factor = best_row["twist factor"]
        rotor_rpm    = best_row["rotor rpm"]

        tex = 583 / count
        twist_tpm = twist_factor * 95 / math.sqrt(tex) * 10
        take_up_speed = rotor_rpm / twist_tpm
        spindle_prod_kg_h = take_up_speed * 60 * tex / 1_000_000
        rotor_prod_day = spindle_prod_kg_h * 24
        working_eff_per_spindle_day = rotor_prod_day * 0.9  # 90%

        def line_kg_per_hour(line_name: str) -> float:
            cfg = LINE_CONFIG[line_name]
            total_spindles = cfg["machines"] * cfg["spindles_per_machine"]
            kg_per_day_line = working_eff_per_spindle_day * total_spindles
            return kg_per_day_line / 24.0

        # pool selection (small defaults to Lines 4–5; may expand to all lines if all work is small — handled later)
        pool = LINES_SMALL if 500 <= qty_required <= 2000 else LINES_MAIN

        speeds = {ln: line_kg_per_hour(ln) for ln in pool}
        best_line = max(speeds, key=speeds.get)
        kg_per_hour = speeds[best_line]
        if kg_per_hour <= 0:
            return None, "Calculated zero throughput"

        hours = qty_required / kg_per_hour
        return round_up(hours), None
    except Exception as e:
        return None, f"Error: {str(e)}"

def get_next_best_color(current_color, available_colors, processed_colors):
    if not current_color or current_color not in NEAREST_FAMILIES:
        return available_colors[0] if available_colors else None
    for nearest in NEAREST_FAMILIES.get(current_color, []):
        if nearest in available_colors and nearest not in processed_colors:
            return nearest
    return available_colors[0] if available_colors else None

def sequence_colors_smartly(badges_df):
    unique_colors = badges_df["color_family_norm"].unique().tolist()
    if len(unique_colors) <= 1:
        return badges_df
    color_priority = badges_df.groupby("color_family_norm")["_due_sort"].min().sort_values()
    current_color = color_priority.index[0]
    color_sequence = [current_color]
    remaining_colors = [c for c in unique_colors if c != current_color]
    while remaining_colors:
        next_color = get_next_best_color(current_color, remaining_colors, color_sequence)
        if next_color:
            color_sequence.append(next_color)
            remaining_colors.remove(next_color)
            current_color = next_color
        else:
            color_sequence.append(remaining_colors[0])
            current_color = remaining_colors[0]
            remaining_colors.pop(0)
    badges_df["color_order"] = badges_df["color_family_norm"].map(
        {color: idx for idx, color in enumerate(color_sequence)}
    )
    return badges_df.sort_values(
        by=["color_order", "_due_sort", "required_qty"],
        ascending=[True, True, False]
    ).drop(columns=["color_order"])

# ========================================
# ROBUST ORDER LOADER (handles Proforma sheet)
# ========================================
def load_customer_orders(customer_file) -> (pd.DataFrame, str):
    try:
        xls = pd.ExcelFile(customer_file)
    except Exception as e:
        return None, f"Error opening Excel file: {e}"

    selected_df, best_hit = None, -1
    for sheet in xls.sheet_names:
        try:
            raw = pd.read_excel(customer_file, sheet_name=sheet, header=None)
            hdr_row = _detect_header_row(raw)
            df = pd.read_excel(customer_file, sheet_name=sheet, header=hdr_row)
            df = df.dropna(how="all")
            df = _standardize_columns(df)
            hit = sum(1 for c in REQUIRED_STD_COLS if c in df.columns)
            if hit > best_hit:
                best_hit = hit
                selected_df = df
        except Exception:
            continue

    if selected_df is None or best_hit < 3:
        return None, "Could not detect a valid data table with required columns."

    if "ColorFamilyName" not in selected_df.columns and "Color" in selected_df.columns:
        selected_df["ColorFamilyName"] = selected_df["Color"]

    if "Quantity" in selected_df.columns:
        selected_df["Quantity"] = pd.to_numeric(selected_df["Quantity"], errors="coerce").fillna(0.0)
    else:
        return None, "Missing Quantity column after normalization."

    if "Composition" in selected_df.columns:
        selected_df["Composition"] = selected_df["Composition"].apply(_clean_text)

    if "Due Date" in selected_df.columns:
        selected_df["Due Date"] = selected_df["Due Date"].apply(ensure_date)

    return selected_df, None

# ========================================
# MAIN PROCESSING FUNCTION
# ========================================
@st.cache_data
def process_orders_and_generate_plan(customer_file):
    """Returns (results_dict, not_matched_df, error_msg_or_none)"""
    try:
        machines = pd.read_excel(MACHINE_FILE_PATH)
    except FileNotFoundError:
        return None, None, "Machine configuration file not found at ./reports/machine.xlsx"
    except Exception as e:
        return None, None, f"Error loading machine file: {str(e)}"

    orders, load_err = load_customer_orders(customer_file)
    if load_err:
        return None, None, load_err

    # ── NEW: split "Sample" single orders (0 < qty ≤ 200 kg) out BEFORE matching/scheduling
    samples_df = orders[(orders["Quantity"] > 0) & (orders["Quantity"] <= 200)].copy()
    plan_orders = orders[~orders.index.isin(samples_df.index)].copy()

    matched_results, unmatched_results = [], []

    for _, order in plan_orders.iterrows():
        # guard: skip rows without core fields
        if any(col not in order or pd.isna(order[col]) for col in ["Yarn Count", "Composition", "Yarn Type", "Quantity"]):
            continue

        normalized_count = normalize_count(order["Yarn Count"])
        hours, error = calculate_hours(order, machines)

        if error:
            unmatched_results.append({
                "order_id": order.get("PI NO"),
                "count": normalized_count,
                "blend": order.get("Composition"),
                "yarn_type": order.get("Yarn Type"),
                "color_code": order.get("Color Code"),
                "color_family": order.get("ColorFamilyName"),
                "required_qty": order.get("Quantity", 0.0),
                "reason": error
            })
        else:
            matched_results.append({
                "order_id": order.get("PI NO"),
                "count": normalized_count,
                "blend": order.get("Composition"),
                "yarn_type": order.get("Yarn Type"),
                "color_code": order.get("Color Code"),
                "color_family": order.get("ColorFamilyName"),
                "required_qty": order.get("Quantity", 0.0),
                "calculated_hours": hours
            })

    df_matched = pd.DataFrame(matched_results)
    df_unmatched = pd.DataFrame(unmatched_results)

    # If there is nothing to schedule (except samples), still return usable payload
    if df_matched.empty:
        empty_results = {
            "production_plan": pd.DataFrame(),
            "batch_status": pd.DataFrame(),
            "line_utilization": pd.DataFrame(),
            "color_changeover": pd.DataFrame(),
            "line_color_summary": pd.DataFrame(),
            "not_matched": df_unmatched,
            "total_pi": int(plan_orders["PI NO"].nunique()) if "PI NO" in plan_orders.columns else len(plan_orders),
            "samples": samples_df.reset_index(drop=True)
        }
        return empty_results, df_unmatched, None

    # Build batches ("badges")
    badges = (
        df_matched
        .groupby(["count", "yarn_type", "blend", "color_code", "color_family"], as_index=False)
        .agg({
            "order_id": lambda x: ", ".join(x.astype(str)),
            "required_qty": "sum",
            "calculated_hours": "sum"
        })
    )

    badges["batch_id"] = badges.apply(
        lambda row: f"{int(row['count'])}-{str(row.get('blend') or '').split()[0]}-{str(row.get('color_code') or '')}".replace(' ', '_'),
        axis=1
    )

    badges["color_family_norm"] = badges["color_family"].fillna("Unknown").astype(str).str.strip().str.title()
    badges["_due_sort"] = datetime.max.date()
    badges["earliest_due"] = None
    badges = sequence_colors_smartly(badges)

    plan_start = datetime.now().date()
    horizon_end = plan_start + timedelta(days=HORIZON_DAYS)

    # Are ALL planned batches small (200–2000)?
    all_small = bool(len(badges) > 0 and (badges["required_qty"].between(200, 2000, inclusive="both").all()))

    # capacity[(date, line, shift_idx)] = remaining_kg for that slot
    capacity = {}
    slot_used = {}

    # separate color→line maps for main vs small pools to preserve color stability
    color_line_map_main  = {}
    color_line_map_small = {}

    def get_or_assign_line_for_pool(color_family, pool):
        cmap = color_line_map_small if pool is LINES_SMALL else color_line_map_main
        if color_family in cmap:
            return cmap[color_family]
        idx = len(cmap) % len(pool)
        ln = pool[idx]
        cmap[color_family] = ln
        return ln

    alloc_rows = []

    for _, badge in badges.iterrows():
        color_family = badge["color_family_norm"]
        total_required = float(badge["required_qty"])
        batch_id = badge["batch_id"]

        # Decide pool:
        # - if ALL batches are small -> allow any line (maximize utilization)
        # - else: small (200–2000) -> Lines 4–5; big -> Lines 1–3
        if all_small:
            target_pool = LINES
        else:
            target_pool = LINES_SMALL if 200 <= total_required <= 2000 else LINES_MAIN

        assigned_line = get_or_assign_line_for_pool(color_family, target_pool)

        remaining = total_required
        current_date = plan_start

        while remaining > 1e-6 and current_date <= horizon_end:
            allocated_any = False

            # Try assigned line first; if very big badge, allow spreading within the same pool
            pool_order = [assigned_line] + [l for l in target_pool if l != assigned_line]
            if remaining > sum(LINE_CONFIG[l]["daily_capacity_kg"] for l in target_pool):
                pool_order = target_pool  # spread across pool

            # Pass 1: allocate within target_pool
            for line in pool_order:
                if remaining <= 1e-6:
                    break

                for shift_idx, (shift_name, shift_start, shift_end) in enumerate(SHIFTS):
                    if remaining <= 1e-6:
                        break

                    key = (current_date, line, shift_idx)
                    if key not in capacity:
                        capacity[key] = PER_SHIFT_CAPACITY[line]

                    avail = capacity[key]
                    if avail <= 1e-9:
                        continue

                    used = min(avail, remaining)
                    capacity[key] -= used

                    used_before = slot_used.get(key, 0.0)
                    slot_used[key] = used_before + used

                    per_shift_cap = PER_SHIFT_CAPACITY[line]

                    shift_day_start = datetime.combine(current_date, datetime.min.time()) + timedelta(minutes=shift_start)
                    start_offset_min = (used_before / per_shift_cap) * SHIFT_DURATION_MIN
                    duration_min = (used / per_shift_cap) * SHIFT_DURATION_MIN

                    start_dt = shift_day_start + timedelta(minutes=start_offset_min)
                    end_dt = start_dt + timedelta(minutes=duration_min)

                    alloc_rows.append({
                        "batch_id": batch_id,
                        "orders": badge["order_id"],
                        "line": line,
                        "date": current_date,
                        "shift": shift_name,
                        "allocated_kg": used,
                        "start_dt": start_dt,
                        "end_dt": end_dt,
                        "color_code": badge.get("color_code"),
                        "color_family": color_family,
                        "count": badge["count"],
                        "blend": badge["blend"],
                        "yarn_type": badge["yarn_type"]
                    })

                    remaining -= used
                    allocated_any = True

            # Pass 2 (NEW): if badge is small and we’re not in all_small mode,
            # and Lines 4–5 couldn't finish it today, try big lines today to fill spare capacity.
            if (remaining > 1e-6) and (not all_small) and (200 <= total_required <= 2000):
                overflow_lines = [l for l in LINES_MAIN if l not in target_pool]  # big lines
                for line in overflow_lines:
                    if remaining <= 1e-6:
                        break
                    for shift_idx, (shift_name, shift_start, shift_end) in enumerate(SHIFTS):
                        if remaining <= 1e-6:
                            break
                        key = (current_date, line, shift_idx)
                        if key not in capacity:
                            capacity[key] = PER_SHIFT_CAPACITY[line]
                        avail = capacity[key]
                        if avail <= 1e-9:
                            continue

                        used = min(avail, remaining)
                        capacity[key] -= used

                        used_before = slot_used.get(key, 0.0)
                        slot_used[key] = used_before + used

                        per_shift_cap = PER_SHIFT_CAPACITY[line]
                        shift_day_start = datetime.combine(current_date, datetime.min.time()) + timedelta(minutes=shift_start)
                        start_offset_min = (used_before / per_shift_cap) * SHIFT_DURATION_MIN
                        duration_min = (used / per_shift_cap) * SHIFT_DURATION_MIN

                        start_dt = shift_day_start + timedelta(minutes=start_offset_min)
                        end_dt = start_dt + timedelta(minutes=duration_min)

                        alloc_rows.append({
                            "batch_id": batch_id,
                            "orders": badge["order_id"],
                            "line": line,
                            "date": current_date,
                            "shift": shift_name,
                            "allocated_kg": used,
                            "start_dt": start_dt,
                            "end_dt": end_dt,
                            "color_code": badge.get("color_code"),
                            "color_family": color_family,
                            "count": badge["count"],
                            "blend": badge["blend"],
                            "yarn_type": badge["yarn_type"]
                        })

                        remaining -= used
                        allocated_any = True

            if not allocated_any:
                current_date += timedelta(days=1)

    df_alloc = pd.DataFrame(alloc_rows)

    # Build results, even if no allocations (only samples)
    if df_alloc.empty:
        results = {
            "production_plan": pd.DataFrame(),
            "batch_status": pd.DataFrame(),
            "line_utilization": pd.DataFrame(),
            "color_changeover": pd.DataFrame(),
            "line_color_summary": pd.DataFrame(),
            "not_matched": df_unmatched,
            "total_pi": int(plan_orders["PI NO"].nunique()) if "PI NO" in plan_orders.columns else len(plan_orders),
            "samples": samples_df.reset_index(drop=True)
        }
        return results, df_unmatched, None

    df_alloc["date"] = pd.to_datetime(df_alloc["date"])

    # Line-specific hours
    df_alloc["no_of_hours"] = df_alloc.apply(
        lambda r: (r["allocated_kg"] / PER_SHIFT_CAPACITY[r["line"]]) * (SHIFT_DURATION_MIN / 60.0),
        axis=1
    )

    # Batch status
    badge_status = []
    for batch_id, g in df_alloc.groupby("batch_id"):
        # Get first row from badges for metadata
        br = badges[badges["batch_id"] == batch_id].iloc[0]
        badge_status.append({
            "batch_id": batch_id,
            "orders": br["order_id"],
            "count": br["count"],
            "blend": br["blend"],
            "yarn_type": br["yarn_type"],
            "color_code": br.get("color_code"),
            "color_family": br.get("color_family"),
            "total_qty": round(g["allocated_kg"].sum(), 2),
            "completion_dt": g["end_dt"].max(),
            "hours_taken": round(g["no_of_hours"].sum(), 3)
        })
    df_badge_status = pd.DataFrame(badge_status)

    # Line utilization (per-line daily capacity)
    util_rows = []
    for line in LINES:
        day_cap = LINE_CONFIG[line]["daily_capacity_kg"]
        for d in sorted(df_alloc["date"].dt.date.unique()):
            used = df_alloc[(df_alloc["line"] == line) & (df_alloc["date"].dt.date == d)]["allocated_kg"].sum()
            util_rows.append({
                "line": line,
                "date": d,
                "capacity_kg": day_cap,
                "used_kg": round(used, 2),
                "util_pct": round((used / day_cap) * 100, 2) if day_cap > 0 else 0.0
            })
    df_line_util = pd.DataFrame(util_rows)

    # Color changeover log
    color_changes = []
    for line in LINES:
        line_alloc = df_alloc[df_alloc["line"] == line].sort_values(["date", "shift"])
        prev_color = None
        for _, row in line_alloc.iterrows():
            if prev_color and prev_color != row["color_family"]:
                color_changes.append({
                    "line": line,
                    "date": row["date"].date() if hasattr(row["date"], "date") else row["date"],
                    "shift": row["shift"],
                    "from_color": prev_color,
                    "to_color": row["color_family"]
                })
            prev_color = row["color_family"]
    df_color_changes = pd.DataFrame(color_changes)

    # Line color summary
    line_color_summary = []
    for line in LINES:
        line_data = df_alloc[df_alloc["line"] == line]
        if not line_data.empty:
            total_line = line_data["allocated_kg"].sum()
            for color, qty in line_data.groupby("color_family")["allocated_kg"].sum().items():
                line_color_summary.append({
                    "line": line,
                    "color_family": color,
                    "total_kg": round(qty, 2),
                    "percentage": round((qty / total_line) * 100, 2) if total_line > 0 else 0.0
                })
    df_line_color_summary = pd.DataFrame(line_color_summary)

    total_pi = int(plan_orders["PI NO"].nunique()) if "PI NO" in plan_orders.columns else len(plan_orders)

    results = {
        "production_plan": df_alloc,
        "batch_status": df_badge_status,
        "line_utilization": df_line_util,
        "color_changeover": df_color_changes,
        "line_color_summary": df_line_color_summary,
        "not_matched": df_unmatched,
        "total_pi": total_pi,
        "samples": samples_df.reset_index(drop=True)  # NEW
    }
    return results, df_unmatched, None

# ========================================
# STREAMLIT UI
# ========================================
st.markdown("""
<div class="hero-header">
    <h1 class="hero-title">Cyclo Production Planning AI</h1>
    <p class="hero-sub">Intelligent Color-Optimized Spinning Mill Production Scheduler</p>
</div>
""", unsafe_allow_html=True)

col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.markdown("""
    <div class="upload-card">
        <div style="font-size:2.2rem">📁</div>
        <div style="font-size:1.2rem;font-weight:600;margin-top:0.5rem">Upload Customer Order File</div>
        <div style="color:#666;margin-top:0.4rem">Excel file (any sheet); headers will be auto-detected</div>
    </div>
    """, unsafe_allow_html=True)

customer_file = st.file_uploader("Choose file", type=["xlsx", "xls"], label_visibility="collapsed")

if customer_file:
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1.5, 1, 1.5])
    with col2:
        generate_btn = st.button("🚀 Generate Plan", type="primary", use_container_width=True)

    if generate_btn:
        with st.spinner("🔄 Processing orders and optimizing schedule..."):
            results, not_matched_df, error = process_orders_and_generate_plan(customer_file)

        if error:
            st.error(f"❌ {error}")
        elif results:
            prod_df = results['production_plan']
            if not prod_df.empty:
                period_start = prod_df['date'].dt.date.min()
                period_end = prod_df['date'].dt.date.max()
            else:
                period_start = datetime.now().date()
                period_end = period_start + timedelta(days=HORIZON_DAYS)

            def fmt(d): return f"{d.day}/{d.month}/{d.year}"
            period_str = f"{fmt(period_start)} - {fmt(period_end)}"
            st.success(f"✅ Production plan for period {period_str} generated successfully!")
            st.markdown("<br>", unsafe_allow_html=True)

            # Top metrics
            total_qty = results['production_plan']['allocated_kg'].sum() if not results['production_plan'].empty else 0
            total_pi = results.get('total_pi', 0)
            total_batch = len(results['batch_status'])
            avg_util = results['line_utilization']['util_pct'].mean() if not results['line_utilization'].empty else 0.0
            sample_count = len(results.get("samples", pd.DataFrame()))
            sample_qty = results.get("samples", pd.DataFrame()).get("Quantity", pd.Series(dtype=float)).sum() if sample_count else 0

            c1, c2, c3, c4, c5 = st.columns(5)
            with c1: st.markdown(f"""<div class="metric-card"><div class="metric-label">Total Quantity (KG)</div><div class="metric-value">{total_qty:,.0f}</div></div>""", unsafe_allow_html=True)
            with c2: st.markdown(f"""<div class="metric-card"><div class="metric-label">Total PI</div><div class="metric-value">{total_pi}</div></div>""", unsafe_allow_html=True)
            with c3: st.markdown(f"""<div class="metric-card"><div class="metric-label">Total Batch</div><div class="metric-value">{total_batch}</div></div>""", unsafe_allow_html=True)
            with c4: st.markdown(f"""<div class="metric-card"><div class="metric-label">Avg Utilization</div><div class="metric-value">{avg_util:.1f}%</div></div>""", unsafe_allow_html=True)
            with c5: st.markdown(f"""<div class="metric-card"><div class="metric-label">Samples (Orders)</div><div class="metric-value">{sample_count} ({sample_qty:,.0f} kg)</div></div>""", unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            st.markdown('<div class="section-header">📈 Visual Analytics</div>', unsafe_allow_html=True)
            if not results['line_utilization'].empty:
                fig_util = px.bar(
                    results['line_utilization'].groupby('line')['util_pct'].mean().reset_index(),
                    x='line', y='util_pct',
                    title='Average Line Utilization (%)',
                    color='util_pct',
                    color_continuous_scale=['#ff6b6b', '#ffd93d', '#63913A'],
                    labels={'util_pct': 'Utilization %', 'line': 'Production Line'}
                )
                fig_util.update_layout(height=350, showlegend=False, plot_bgcolor='white', paper_bgcolor='white')
                st.plotly_chart(fig_util, use_container_width=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="section-header">📋 Industrial Engineering Plan</div>', unsafe_allow_html=True)

            tabs = st.tabs([
                "📅 Production Schedule",
                "📊 Batch Summary",
                "📈 Line Utilization",
                "🎨 Color Changeover",
                "🎯 Color Distribution",
                "🧪 Samples",
                "⚠️ Exceptions"
            ])

            with tabs[0]:
                st.markdown("<br>", unsafe_allow_html=True)
                schedule_df = results['production_plan'].copy()
                if not schedule_df.empty:
                    schedule_df['date'] = pd.to_datetime(schedule_df['date']).dt.date
                    front = ['orders', 'line', 'date', 'shift', 'allocated_kg']
                    rest = [c for c in schedule_df.columns if c not in front]
                    schedule_df = schedule_df.reindex(columns=front+rest)
                    st.dataframe(schedule_df.reset_index(drop=True), use_container_width=True, height=450)
                else:
                    st.info("No production schedule available.")

            with tabs[1]:
                st.markdown("<br>", unsafe_allow_html=True)
                batch_df = results['batch_status'].copy()
                if not batch_df.empty:
                    front = ['yarn_type', 'orders', 'count', 'blend']
                    rest = [c for c in batch_df.columns if c not in front]
                    batch_df = batch_df.reindex(columns=[c for c in front if c in batch_df.columns] + rest)
                    batch_df = batch_df.sort_values(['yarn_type'])
                    st.dataframe(batch_df.reset_index(drop=True), use_container_width=True, height=450)
                else:
                    st.info("No batch summary available.")

            with tabs[2]:
                st.markdown("<br>", unsafe_allow_html=True)
                if not results['line_utilization'].empty:
                    util_display = results['line_utilization'].copy()
                    util_display['date'] = pd.to_datetime(util_display['date']).dt.date
                    st.dataframe(util_display.reset_index(drop=True), use_container_width=True, height=350)
                    c1, c2, c3 = st.columns(3)
                    with c1: st.metric("Peak Utilization", f"{util_display['util_pct'].max():.1f}%")
                    with c2: st.metric("Minimum Utilization", f"{util_display['util_pct'].min():.1f}%")
                    with c3: st.metric("Total Capacity", f"{util_display['capacity_kg'].sum():,.0f} kg")
                else:
                    st.info("No line utilization data available.")

            with tabs[3]:
                st.markdown("<br>", unsafe_allow_html=True)
                cc = results['color_changeover'].copy()
                if not cc.empty:
                    if 'date' in cc.columns:
                        cc['date'] = pd.to_datetime(cc['date']).dt.date
                    st.dataframe(cc.reset_index(drop=True), use_container_width=True, height=450)
                    st.info(f"Total color changeovers: {len(cc)}")
                else:
                    st.success("🎉 No color changeovers — optimal batching achieved.")

            with tabs[4]:
                st.markdown("<br>", unsafe_allow_html=True)
                lcd = results['line_color_summary'].copy()
                if not lcd.empty:
                    st.dataframe(lcd.sort_values(['line','total_kg'], ascending=[True, False]).reset_index(drop=True), use_container_width=True, height=450)
                    for line in LINES:
                        line_data = lcd[lcd['line'] == line]
                        if not line_data.empty:
                            st.markdown(f"**{line}:**")
                            cols = st.columns(len(line_data))
                            for idx, (_, row) in enumerate(line_data.iterrows()):
                                with cols[idx]:
                                    st.metric(row['color_family'], f"{row['total_kg']:,.0f} kg", f"{row['percentage']:.1f}%")
                else:
                    st.info("No color distribution data available.")

            with tabs[5]:
                st.markdown("<br>", unsafe_allow_html=True)
                samples_view = results.get("samples", pd.DataFrame()).copy()
                if not samples_view.empty:
                    # Keep common columns in a friendly order if present
                    desired = ["PI NO","Customer","Invoice Date","Yarn Count","Composition","Yarn Type","Color Code","ColorFamilyName","Quantity","Due Date","Color"]
                    cols = [c for c in desired if c in samples_view.columns] + [c for c in samples_view.columns if c not in desired]
                    st.dataframe(samples_view[cols].reset_index(drop=True), use_container_width=True, height=450)
                    st.info(f"{len(samples_view)} sample orders moved to 'Sample' sheet.")
                else:
                    st.success("No sample orders (≤ 200 kg).")

            with tabs[6]:
                st.markdown("<br>", unsafe_allow_html=True)
                not_matched = results.get('not_matched', pd.DataFrame())
                if not not_matched.empty:
                    st.dataframe(not_matched.reset_index(drop=True), use_container_width=True, height=450)
                    st.warning(f"{len(not_matched)} orders could not be matched. Check 'reason' column for details.")
                else:
                    st.success("No exceptions — all orders matched successfully.")

            st.markdown("<br><br>", unsafe_allow_html=True)
            st.markdown('<div class="section-header">📥 Download Production Plan</div>', unsafe_allow_html=True)

            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.markdown("""
                <div style="padding:1rem;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.06)">
                    <h3 style="color:#63913A;margin-top:0">Complete Production Plan Package</h3>
                    <p>Includes: Production Schedule, Batch Summary, Line Utilization, Color Changeover, Color Distribution, <strong>Sample Orders</strong>, and Exceptions.</p>
                </div>
                """, unsafe_allow_html=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    prod_export = results['production_plan'].copy()
                    if not prod_export.empty:
                        prod_export['date'] = pd.to_datetime(prod_export['date']).dt.date
                        prod_export.to_excel(writer, sheet_name="ProductionSchedule", index=False)

                    batch_export = results['batch_status'].copy()
                    if not batch_export.empty:
                        batch_export.to_excel(writer, sheet_name="BatchSummary", index=False)

                    if not results['line_utilization'].empty:
                        lu = results['line_utilization'].copy()
                        lu['date'] = pd.to_datetime(lu['date']).dt.date
                        lu.to_excel(writer, sheet_name="LineUtilization", index=False)

                    if not results['color_changeover'].empty:
                        cc = results['color_changeover'].copy()
                        if 'date' in cc.columns:
                            cc['date'] = pd.to_datetime(cc['date']).dt.date
                        cc.to_excel(writer, sheet_name="ColorChangeover", index=False)

                    if not results['line_color_summary'].empty:
                        results['line_color_summary'].to_excel(writer, sheet_name="LineColorDistribution", index=False)

                    # NEW: write Sample sheet always (even if empty, we’ll skip)
                    samples_export = results.get("samples", pd.DataFrame()).copy()
                    if not samples_export.empty:
                        samples_export.to_excel(writer, sheet_name="Sample", index=False)

                    if not results['not_matched'].empty:
                        results['not_matched'].to_excel(writer, sheet_name="NotMatchedOrders", index=False)

                output.seek(0)
                st.download_button(
                    label="📥 Download Complete Production Plan (Excel)",
                    data=output,
                    file_name=f"cyclo_production_plan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
else:
    st.markdown("<br><br>", unsafe_allow_html=True)
    l, r = st.columns(2)
    with l:
        st.markdown("""
        <div style="padding:1rem;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.06)">
            <h3 style="color:#63913A;margin-top:0">🎯 Key Features</h3>
            <ul>
                <li><strong>Smart Color Sequencing</strong> to minimize changeovers</li>
                <li><strong>Optimal Line Assignment</strong> by color family</li>
                <li><strong>Real-time Scheduling</strong> across shifts & days</li>
                <li><strong>Capacity Planning</strong> for all production lines</li>
                <li><strong>Visual Analytics</strong> & complete Excel export (incl. <em>Sample</em> sheet)</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    with r:
        st.markdown("""
        <div style="padding:1rem;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.06)">
            <h3 style="color:#63913A;margin-top:0">📋 File Requirements</h3>
            <p>Upload any Excel sheet; the app auto-detects the header row.</p>
            <p><strong>Must include columns (any label variant):</strong></p>
            <ul>
                <li>PI No</li><li>Yarn Count</li><li>Composition</li><li>Yarn Type</li><li>Quantity (kg)</li>
                <li>Optional: Color Code, Color Family/Color, Due Date</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
