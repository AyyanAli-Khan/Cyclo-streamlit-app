import streamlit as st
import pandas as pd
import math
import io
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go

# ========================================
# PAGE CONFIGURATION
# ========================================
st.set_page_config(
    page_title="Cyclo Production Planning AI",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ========================================
# STYLES (unchanged except minor header area)
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
    .hero-title { font-size: 1.8rem; font-weight:700; margin:0; }
    .hero-sub { margin:0.25rem 0 0 0; opacity:0.95; }
    .upload-card { padding: 2rem; border-radius: 12px; border: 2px solid #e8f5e9; box-shadow: 0 4px 6px rgba(0,0,0,0.07); text-align:center; transition: all 0.3s ease; }
    .upload-card:hover { border-color: #63913A; box-shadow: 0 8px 20px rgba(99,145,58,0.2); }

        /* Buttons */
    .stButton > button {
        background: linear-gradient(135deg, #63913A 0%, #7AB850 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 1rem 3rem;
        font-weight: 600;
        font-size: 1.1rem;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(99, 145, 58, 0.3);
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(99, 145, 58, 0.4);
    }
    

    .metric-container { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1.2rem; margin: 1.5rem 0; }
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
BLEND_MAPPING = {
    "70% CYCLO® Recycled Cotton 30% Recycled Polyester": "70/30 CYL Cot/poly",
    "70% CYCLO® Recycled Cotton 30% Polyester": "70/30 CYL Cot/poly",
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
}

NEAREST_FAMILIES = {
    "RED": ["Maroon", "Rust Melange", "Orange"],
    "Green": ["Midnight Olive", "Pearl Teal"],
    "Blue": ["Denim", "Midnight Blue", "Pearl Teal"],
    "Beige": ["Stone", "Cream", "Natural"],
    "White": ["Cream", "Grey"],
    "Yellow": ["Dijon", "Cream"],
    "Stone": ["Beige", "Natural", "Grey"],
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

CAPACITY_PER_LINE_PER_DAY = 5000.0
LINES = ["Line 1", "Line 2", "Line 3"]
SHIFTS = [("A", 0, 480), ("B", 480, 960), ("C", 960, 1440)]
SHIFT_DURATION_MIN = 480
PER_SHIFT_CAPACITY = CAPACITY_PER_LINE_PER_DAY / len(SHIFTS)
HORIZON_DAYS = 60
MACHINE_FILE_PATH = "./reports/machine.xlsx"

# ========================================
# HELPER FUNCTIONS
# ========================================
def round_up(val: float) -> float:
    return math.ceil(val * 100) / 100

def normalize_count(count_value):
    if isinstance(count_value, str) and "/" in count_value:
        return int(count_value.split("/")[0].strip())
    return int(count_value)

def normalize_blend(blend_raw: str) -> str:
    clean_blend = " ".join(str(blend_raw).split())
    return BLEND_MAPPING.get(clean_blend, None)

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
    try:
        count = normalize_count(order["Yarn Count"])
        blend_raw = str(order["Composition"]).strip()
        yarn_type = str(order["Yarn Type"]).strip()
        qty_required = order["Quantity"]

        blend = normalize_blend(blend_raw)
        if not blend:
            return None, f"Blend not mapped: {blend_raw}"

        matched = machines[
            (machines["Counts"] == count) &
            (machines["Blends"] == blend) &
            (machines["Yarn Type"] == yarn_type)
        ]

        if matched.empty:
            return None, f"No machine data"

        best_row = matched.loc[matched["twist factor"].idxmax()]
        twist_factor = best_row["twist factor"]
        rotor_rpm = best_row["rotor rpm"]

        tex = 583 / count
        twist_tpm = twist_factor * 95 / math.sqrt(tex) * 10
        take_up_speed = rotor_rpm / twist_tpm
        spindle_prod_kg_h = take_up_speed * 60 * tex / 1_000_000
        rotor_prod_day = spindle_prod_kg_h * 24
        working_eff = rotor_prod_day * 0.9
        required_days = qty_required / working_eff
        per_machine_spindles = 460
        days = required_days / per_machine_spindles
        per_line_machines = 3
        total_days = days / per_line_machines
        total_hours = total_days * 24

        return round_up(total_hours), None
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

    try:
        orders = pd.read_excel(customer_file, sheet_name="Sheet1")
    except Exception as e:
        return None, None, f"Error reading customer file: {str(e)}"

    matched_results = []
    unmatched_results = []

    for _, order in orders.iterrows():
        normalized_count = normalize_count(order["Yarn Count"])
        hours, error = calculate_hours(order, machines)

        if error:
            unmatched_results.append({
                "order_id": order["PI NO"],
                "count": normalized_count,
                "blend": order["Composition"],
                "yarn_type": order["Yarn Type"],
                "color_code": order.get("Color Code", None),
                "color_family": order.get("ColorFamilyName", None),
                "required_qty": order["Quantity"],
                "reason": error
            })
        else:
            matched_results.append({
                "order_id": order["PI NO"],
                "count": normalized_count,
                "blend": order["Composition"],
                "yarn_type": order["Yarn Type"],
                "color_code": order.get("Color Code", None),
                "color_family": order.get("ColorFamilyName", None),
                "required_qty": order["Quantity"],
                "calculated_hours": hours
            })

    df_matched = pd.DataFrame(matched_results)
    df_unmatched = pd.DataFrame(unmatched_results)

    if df_matched.empty:
        # still return unmatched so UI can show exceptions
        return None, df_unmatched, "No matched orders found"

    badges = (
        df_matched
        .groupby(["count", "yarn_type", "blend", "color_code", "color_family"], as_index=False)
        .agg({
            "order_id": lambda x: ", ".join(x.astype(str)),
            "required_qty": "sum",
            "calculated_hours": "sum"
        })
    )

    # create batch (badge) id for internal grouping
    badges["batch_id"] = badges.apply(
        lambda row: f"{int(row['count'])}-{str(row.get('blend') or '').split()[0]}-{str(row.get('color_code') or '')}".replace(" ", "_"),
        axis=1
    )

    badges["color_family_norm"] = badges["color_family"].fillna("Unknown").astype(str).str.strip().str.title()
    badges["_due_sort"] = datetime.max.date()
    badges["earliest_due"] = None
    badges = sequence_colors_smartly(badges)

    plan_start = datetime.now().date()
    horizon_end = plan_start + timedelta(days=HORIZON_DAYS)

    capacity = {}
    slot_used = {}
    color_family_line_map = {}
    alloc_rows = []

    def get_or_assign_line(color_family):
        if color_family in color_family_line_map:
            return color_family_line_map[color_family]
        assigned_line = LINES[len(color_family_line_map) % len(LINES)]
        color_family_line_map[color_family] = assigned_line
        return assigned_line

    current_color = None
    assigned_line = None

    for _, badge in badges.iterrows():
        color_family = badge["color_family_norm"]

        if color_family != current_color:
            current_color = color_family
            assigned_line = get_or_assign_line(color_family)

        remaining = float(badge["required_qty"])
        batch_id = badge["batch_id"]
        current_date = plan_start

        while remaining > 1e-6 and current_date <= horizon_end:
            allocated = False
            lines_to_try = [assigned_line]

            if remaining > CAPACITY_PER_LINE_PER_DAY * 3:
                lines_to_try.extend([l for l in LINES if l != assigned_line])

            for line in lines_to_try:
                if remaining <= 1e-6:
                    break

                for shift_idx, (shift_name, shift_start, shift_end) in enumerate(SHIFTS):
                    if remaining <= 1e-6:
                        break

                    key = (current_date, line, shift_idx)
                    if key not in capacity:
                        capacity[key] = PER_SHIFT_CAPACITY

                    avail = capacity[key]
                    if avail <= 1e-9:
                        continue

                    used = min(avail, remaining)
                    capacity[key] -= used

                    used_before = slot_used.get(key, 0.0)
                    slot_used[key] = used_before + used

                    shift_day_start = datetime.combine(current_date, datetime.min.time()) + timedelta(minutes=shift_start)
                    start_offset = (used_before / PER_SHIFT_CAPACITY) * SHIFT_DURATION_MIN
                    duration = (used / PER_SHIFT_CAPACITY) * SHIFT_DURATION_MIN
                    start_dt = shift_day_start + timedelta(minutes=start_offset)
                    end_dt = start_dt + timedelta(minutes=duration)

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
                    allocated = True

            if not allocated:
                current_date += timedelta(days=1)

    df_alloc = pd.DataFrame(alloc_rows)
    df_alloc["date"] = pd.to_datetime(df_alloc["date"])
    df_alloc["no_of_hours"] = df_alloc["allocated_kg"] * 8.0 / PER_SHIFT_CAPACITY

    # Batch (badge) status summary
    badge_status = []
    for batch_id, g in df_alloc.groupby("batch_id"):
        badge_row = badges[badges["batch_id"] == batch_id].iloc[0]
        badge_status.append({
            "batch_id": batch_id,
            "orders": badge_row["order_id"],
            "count": badge_row["count"],
            "blend": badge_row["blend"],
            "yarn_type": badge_row["yarn_type"],
            "color_code": badge_row.get("color_code"),
            "color_family": badge_row.get("color_family"),
            "total_qty": g["allocated_kg"].sum(),
            "completion_dt": g["end_dt"].max(),
            "hours_taken": round((g["allocated_kg"] * 8.0 / PER_SHIFT_CAPACITY).sum(), 3)
        })
    df_badge_status = pd.DataFrame(badge_status)

    util_rows = []
    for line in LINES:
        for d in sorted(df_alloc["date"].dt.date.unique()):
            used = df_alloc[(df_alloc["line"] == line) & (df_alloc["date"].dt.date == d)]["allocated_kg"].sum()
            util_rows.append({
                "line": line,
                "date": d,
                "capacity_kg": CAPACITY_PER_LINE_PER_DAY,
                "used_kg": used,
                "util_pct": round((used / CAPACITY_PER_LINE_PER_DAY) * 100, 2)
            })
    df_line_util = pd.DataFrame(util_rows)

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

    line_color_summary = []
    for line in LINES:
        line_data = df_alloc[df_alloc["line"] == line]
        if not line_data.empty:
            for color, qty in line_data.groupby("color_family")["allocated_kg"].sum().items():
                line_color_summary.append({
                    "line": line,
                    "color_family": color,
                    "total_kg": round(qty, 2),
                    "percentage": round((qty / line_data["allocated_kg"].sum()) * 100, 2)
                })
    df_line_color_summary = pd.DataFrame(line_color_summary)

    # total PI count from original orders
    total_pi = int(orders["PI NO"].nunique()) if "PI NO" in orders.columns else len(orders)

    results = {
        "production_plan": df_alloc,
        "batch_status": df_badge_status,
        "line_utilization": df_line_util,
        "color_changeover": df_color_changes,
        "line_color_summary": df_line_color_summary,
        "not_matched": df_unmatched,
        "total_pi": total_pi
    }

    return results, df_unmatched, None

# ========================================
# STREAMLIT UI
# ========================================

st.markdown("""
<div class="hero-header">
    <h1>Cyclo Production Planning AI</h1>
    <p>Intelligent Color-Optimized Spinning Mill Production Scheduler</p>
</div>
""", unsafe_allow_html=True)

# Upload Section
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.markdown("""
    <div class="upload-card">
        <div style="font-size:2.2rem">📁</div>
        <div style="font-size:1.2rem;font-weight:600;margin-top:0.5rem">Upload Customer Order File</div>
        <div style="color:#666;margin-top:0.4rem">Excel file with order details (Sheet1)</div>
    </div>
    """, unsafe_allow_html=True)

customer_file = st.file_uploader(
    "Choose file",
    type=["xlsx", "xls"],
    label_visibility="collapsed"
)

# Process Section
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
            # compute period from production_plan dates
            prod_df = results['production_plan']
            if not prod_df.empty:
                period_start = prod_df['date'].dt.date.min()
                period_end = prod_df['date'].dt.date.max()
            else:
                period_start = datetime.now().date()
                period_end = period_start + timedelta(days=HORIZON_DAYS)

            def fmt(d):
                return f"{d.day}/{d.month}/{d.year}"

            period_str = f"{fmt(period_start)} - {fmt(period_end)}"

            st.success(f"✅ Production plan for period {period_str} generated successfully!")

            st.markdown("<br>", unsafe_allow_html=True)

            # Metrics Dashboard (AI Roadmap heading with period)
            st.markdown(f'<div class="section-header">🗺️ AI Roadmap ({period_str})</div>', unsafe_allow_html=True)

            col1, col2, col3, col4 = st.columns(4)

            total_qty = results['production_plan']['allocated_kg'].sum() if not results['production_plan'].empty else 0
            total_pi = results.get('total_pi', 0)
            total_batch = len(results['batch_status'])
            avg_util = results['line_utilization']['util_pct'].mean() if not results['line_utilization'].empty else 0.0

            with col1:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">Total Quantity (KG)</div>
                    <div class="metric-value">{total_qty:,.0f}</div>
                </div>
                """, unsafe_allow_html=True)

            with col2:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">Total PI</div>
                    <div class="metric-value">{total_pi}</div>
                </div>
                """, unsafe_allow_html=True)

            with col3:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">Total Batch</div>
                    <div class="metric-value">{total_batch}</div>
                </div>
                """, unsafe_allow_html=True)

            with col4:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">Avg Utilization</div>
                    <div class="metric-value">{avg_util:.1f}%</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<br><br>", unsafe_allow_html=True)

            # Visualizations Section
            st.markdown('<div class="section-header">📈 Visual Analytics</div>', unsafe_allow_html=True)
            # col1, col2 = st.columns(2)

            # with col1:
            #     if not results['line_utilization'].empty:
            #         fig_util = px.bar(
            #             results['line_utilization'].groupby('line')['util_pct'].mean().reset_index(),
            #             x='line',
            #             y='util_pct',
            #             title='Average Line Utilization (%)',
            #             labels={'util_pct': 'Utilization %', 'line': 'Production Line'}
            #         )
            #         fig_util.update_layout(height=350, showlegend=False, plot_bgcolor='white', paper_bgcolor='white')
            #         st.plotly_chart(fig_util, use_container_width=True)
            #     else:
            #         st.info("No line utilization data available.")
                # Line utilization chart
            fig_util = px.bar(
                results['line_utilization'].groupby('line')['util_pct'].mean().reset_index(),
                x='line',
                y='util_pct',
                title='Average Line Utilization (%)',
                color='util_pct',
                color_continuous_scale=['#ff6b6b', '#ffd93d', '#63913A'],
                labels={'util_pct': 'Utilization %', 'line': 'Production Line'}
            )
            fig_util.update_layout(
                height=350,
                showlegend=False,
                plot_bgcolor='white',
                paper_bgcolor='white'
            )
            st.plotly_chart(fig_util, use_container_width=True)
        
            st.markdown("<br>", unsafe_allow_html=True)

            # Data Tables Section (Industrial Engineering Plan)
            st.markdown('<div class="section-header">📋 Industrial Engineering Plan</div>', unsafe_allow_html=True)

            tabs = st.tabs([
                "📅 Production Schedule",
                "📊 Batch Summary",
                "📈 Line Utilization",
                "🎨 Color Changeover",
                "🎯 Color Distribution",
                "⚠️ Exceptions"
            ])

            # Tab: Production Schedule
            with tabs[0]:
                st.markdown("<br>", unsafe_allow_html=True)
                schedule_df = results['production_plan'].copy()
                if not schedule_df.empty:
                    # ensure date only
                    schedule_df['date'] = pd.to_datetime(schedule_df['date']).dt.date
                    # reorder columns: orders, line, date, shift, allocated_kg, then others
                    cols = schedule_df.columns.tolist()
                    front = ['orders', 'line', 'date', 'shift', 'allocated_kg']
                    rest = [c for c in cols if c not in front]
                    display_cols = front + rest
                    schedule_df = schedule_df.reindex(columns=display_cols)
                    st.dataframe(schedule_df.reset_index(drop=True), use_container_width=True, height=450)
                else:
                    st.info("No production schedule available.")

            # Tab: Batch Summary
            with tabs[1]:
                st.markdown("<br>", unsafe_allow_html=True)
                batch_df = results['batch_status'].copy()
                if not batch_df.empty:
                    # columns pattern: yarn_type, orders, count, blend, rest...
                    cols = batch_df.columns.tolist()
                    front = ['yarn_type', 'orders', 'count', 'blend']
                    rest = [c for c in cols if c not in front]
                    display_cols = [c for c in front if c in cols] + rest
                    batch_df = batch_df.reindex(columns=display_cols)
                    # sort by yarn_type
                    batch_df = batch_df.sort_values(['yarn_type'])
                    st.dataframe(batch_df.reset_index(drop=True), use_container_width=True, height=450)
                else:
                    st.info("No batch summary available.")

            # Tab: Line Utilization
            with tabs[2]:
                st.markdown("<br>", unsafe_allow_html=True)
                if not results['line_utilization'].empty:
                    util_display = results['line_utilization'].copy()
                    util_display['date'] = pd.to_datetime(util_display['date']).dt.date
                    st.dataframe(util_display.reset_index(drop=True), use_container_width=True, height=350)

                    # Additional utilization stats
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        max_util = results['line_utilization']['util_pct'].max()
                        st.metric("Peak Utilization", f"{max_util:.1f}%")
                    with col2:
                        min_util = results['line_utilization']['util_pct'].min()
                        st.metric("Minimum Utilization", f"{min_util:.1f}%")
                    with col3:
                        total_capacity = results['line_utilization']['capacity_kg'].sum()
                        st.metric("Total Capacity", f"{total_capacity:,.0f} kg")
                else:
                    st.info("No line utilization data available.")

            # Tab: Color Changeover
            with tabs[3]:
                st.markdown("<br>", unsafe_allow_html=True)
                cc = results['color_changeover'].copy()
                if not cc.empty:
                    # ensure date column shows date only when possible
                    if 'date' in cc.columns:
                        cc['date'] = pd.to_datetime(cc['date']).dt.date
                    st.dataframe(cc.reset_index(drop=True), use_container_width=True, height=450)
                    st.info(f"Total color changeovers: {len(cc)}")
                else:
                    st.success("🎉 Perfect scheduling! No color changeovers detected - all colors batched optimally on their assigned lines.")

            # Tab: Color Distribution
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

            # Tab: Exceptions / Not Matched
            with tabs[5]:
                st.markdown("<br>", unsafe_allow_html=True)
                not_matched = results.get('not_matched', pd.DataFrame())
                if not not_matched.empty:
                    st.dataframe(not_matched.reset_index(drop=True), use_container_width=True, height=450)
                    st.warning(f"{len(not_matched)} orders could not be matched. Check 'reason' column for details.")
                else:
                    st.success("No exceptions — all orders matched successfully.")

            st.markdown("<br><br>", unsafe_allow_html=True)

            # Download Section
            st.markdown('<div class="section-header">📥 Download Production Plan</div>', unsafe_allow_html=True)

            col1, col2, col3 = st.columns([1, 2, 1])

            with col2:
                st.markdown("""
                <div style="padding:1rem;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.06)">
                    <h3 style="color:#63913A;margin-top:0">Complete Production Plan Package</h3>
                    <p>Download includes all sheets: Production Schedule, Batch Summary, Line Utilization, Color Changeover, Color Distribution, and Exceptions (if any).</p>
                </div>
                """, unsafe_allow_html=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Production Schedule sheet (date export as date)
                    prod_export = results['production_plan'].copy()
                    if not prod_export.empty:
                        prod_export['date'] = pd.to_datetime(prod_export['date']).dt.date
                        prod_export.to_excel(writer, sheet_name="ProductionSchedule", index=False)
                    # Batch summary
                    batch_export = results['batch_status'].copy()
                    if not batch_export.empty:
                        batch_export.to_excel(writer, sheet_name="BatchSummary", index=False)
                    # Line utilization
                    if not results['line_utilization'].empty:
                        lu = results['line_utilization'].copy()
                        lu['date'] = pd.to_datetime(lu['date']).dt.date
                        lu.to_excel(writer, sheet_name="LineUtilization", index=False)
                    # Color changeover
                    if not results['color_changeover'].empty:
                        cc = results['color_changeover'].copy()
                        if 'date' in cc.columns:
                            cc['date'] = pd.to_datetime(cc['date']).dt.date
                        cc.to_excel(writer, sheet_name="ColorChangeover", index=False)
                    # Color distribution
                    if not results['line_color_summary'].empty:
                        results['line_color_summary'].to_excel(writer, sheet_name="LineColorDistribution", index=False)
                    # Exceptions / not matched
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
    # Information Section when no file uploaded
    st.markdown("<br><br>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <div style="padding:1rem;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.06)">
            <h3 style="color:#63913A;margin-top:0">🎯 Key Features</h3>
            <ul>
                <li><strong>Smart Color Sequencing:</strong> Groups similar colors to minimize line changes</li>
                <li><strong>Optimal Line Assignment:</strong> Dedicates lines to color families for efficiency</li>
                <li><strong>Real-time Scheduling:</strong> Allocates orders across shifts and dates</li>
                <li><strong>Capacity Planning:</strong> Ensures optimal utilization of all production lines</li>
                <li><strong>Visual Analytics:</strong> Interactive charts and comprehensive reports</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div style="padding:1rem;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.06)">
            <h3 style="color:#63913A;margin-top:0">📋 File Requirements</h3>
            <p><strong>Customer Order File (Sheet1) must contain:</strong></p>
            <ul>
                <li>PI NO - Purchase/Order number</li>
                <li>Yarn Count - Count specification</li>
                <li>Composition - Blend composition</li>
                <li>Yarn Type - Type of yarn</li>
                <li>Color Code - Specific color code</li>
                <li>ColorFamilyName - Color family</li>
                <li>Quantity - Order quantity (kg)</li>
                <li>Due Date - Delivery date (optional)</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    with st.expander("ℹ️ How It Works"):
        st.markdown("""
        ### Production Planning Process

        1. **Upload:** Submit your customer order Excel file
        2. **Processing:** System matches orders with machine configurations
        3. **Optimization:** AI groups orders by color families using nearest-family logic
        4. **Scheduling:** Allocates production across lines, shifts, and dates
        5. **Download:** Get comprehensive production plan with all details

        ### Color Optimization Logic

        The system uses intelligent color grouping to minimize line cleaning:
        - Natural colors stay together (Natural → Beige → Cream)
        - Blues group together (Blue → Denim → Midnight Blue)
        - Reds batch together (Red → Maroon → Rust Melange)

        ### Machine Configuration

        Machine settings are pre-loaded from `./reports/machine.xlsx` - no need to upload separately!
        """)
