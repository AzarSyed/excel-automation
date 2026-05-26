"""
Excel Automation & Reporting Tool
Streamlit dashboard — entry point.
Run:  streamlit run main.py
"""
import io
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT))

from automation.cleaner import DataCleaner
from automation.validator import DataValidator
from reports.excel_exporter import ExcelExporter
from utils.helpers import compute_analytics, dataframe_info, format_number, standardize_col_name
from utils.logger import setup_logger

logger = setup_logger("dashboard")


# ── Page configuration ────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Excel Automation Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ───────────────────────────────────────────────────────────────────────

st.markdown(
    """
<style>
/* ── Global app background ─────────────────────────────────────── */
.stApp { background: #f0f5fb; }
[data-testid="stAppViewContainer"] > .main { background: #f0f5fb; }
[data-testid="block-container"] { padding-top: 1.8rem; }

/* ── Sidebar ────────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0d1f33 0%, #12294a 100%);
    border-right: 1px solid #1e3a5c;
}
[data-testid="stSidebar"] * { color: #ccddef; }
[data-testid="stSidebar"] h2 { color: #7fb3d3 !important; }
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiselect label,
[data-testid="stSidebar"] .stCheckbox label,
[data-testid="stSidebar"] .stRadio label { color: #90b8d4 !important; font-size: 0.82rem; }
[data-testid="stSidebar"] .stButton > button {
    background: linear-gradient(135deg,#1e5fa3,#2980b9);
    color: white; border: none; font-weight: 600;
    border-radius: 8px; letter-spacing: 0.02em;
}
[data-testid="stSidebar"] .stButton > button:hover {
    background: linear-gradient(135deg,#2471a3,#3498db);
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(41,128,185,0.4);
}
[data-testid="stSidebar"] hr { border-color: #1e3a5c; }

/* ── Main header band ───────────────────────────────────────────── */
.main-header {
    background: linear-gradient(135deg, #0d1f33 0%, #1a4a7a 60%, #2980b9 100%);
    border-radius: 14px;
    padding: 28px 32px 22px 32px;
    color: white;
    margin-bottom: 22px;
    box-shadow: 0 4px 20px rgba(13,31,51,0.18);
}
.main-header h1 {
    font-size: 2rem; font-weight: 800; margin: 0; color: white !important;
    letter-spacing: -0.01em;
}
.main-header p { margin: 6px 0 0 0; opacity: 0.80; font-size: 0.95rem; }

/* ── Step cards (welcome screen) ────────────────────────────────── */
.step-card {
    background: linear-gradient(160deg, #1a3a5c 0%, #1e6ca0 100%);
    border-radius: 12px;
    padding: 20px 18px;
    color: white;
    height: 100%;
    box-shadow: 0 4px 14px rgba(13,31,51,0.16);
    border-top: 3px solid #5dade2;
}
.step-card .step-num {
    font-size: 0.72rem; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.1em; color: #5dade2; margin-bottom: 6px;
}
.step-card .step-title { font-size: 1rem; font-weight: 700; margin-bottom: 6px; }
.step-card .step-desc  { font-size: 0.82rem; opacity: 0.82; line-height: 1.45; }

/* ── Feature cards (welcome screen) ─────────────────────────────── */
.feat-card {
    background: white;
    border-radius: 12px;
    padding: 22px 20px;
    height: 100%;
    box-shadow: 0 2px 12px rgba(13,31,51,0.08);
    border-top: 4px solid var(--accent);
}
.feat-card .feat-title {
    font-size: 1rem; font-weight: 700; color: #1a3a5c;
    margin-bottom: 12px;
}
.feat-card ul { margin: 0; padding-left: 18px; }
.feat-card li { font-size: 0.84rem; color: #34495e; margin-bottom: 4px; line-height: 1.4; }

/* ── KPI row grid ────────────────────────────────────────────────── */
.kpi-row {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
    gap: 12px;
    margin: 4px 0 16px 0;
}
.kpi-card {
    border-radius: 12px;
    padding: 18px 14px;
    text-align: center;
    color: white;
    min-height: 108px;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    box-shadow: 0 4px 16px rgba(0,0,0,0.18);
}
.kpi-card .kpi-value, .kpi-value {
    display: block; font-size: 2rem; font-weight: 800; line-height: 1.1; color: white;
}
.kpi-card .kpi-label, .kpi-label {
    display: block; font-size: 0.72rem; opacity: 0.88; margin-top: 6px;
    letter-spacing: 0.06em; text-transform: uppercase; color: white;
}
.kpi-card .kpi-delta, .kpi-delta {
    display: block; font-size: 0.68rem; opacity: 0.68; margin-top: 3px; color: white;
}

/* ── Section heading ─────────────────────────────────────────────── */
.section-head {
    font-size: 1.05rem; font-weight: 700;
    color: #1a3a5c;
    background: white;
    border-left: 4px solid #2980b9;
    padding: 8px 14px;
    border-radius: 0 8px 8px 0;
    margin: 20px 0 14px 0;
    box-shadow: 0 1px 6px rgba(13,31,51,0.07);
}

/* ── Tabs ────────────────────────────────────────────────────────── */
[data-testid="stTabs"] [role="tab"] {
    font-weight: 600; color: #4a6a8a; padding: 8px 18px;
}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
    color: #1a3a5c;
    border-bottom: 3px solid #2980b9;
}

/* ── Dataframe ───────────────────────────────────────────────────── */
[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }

/* ── Download buttons ────────────────────────────────────────────── */
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #1a4a7a, #2980b9);
    color: white; border: none; border-radius: 8px;
    font-weight: 600; letter-spacing: 0.02em;
}
[data-testid="stDownloadButton"] > button:hover {
    background: linear-gradient(135deg, #1e5fa3, #3498db);
    box-shadow: 0 4px 12px rgba(41,128,185,0.35);
}

/* ── Divider ─────────────────────────────────────────────────────── */
hr { border-color: #d0dcea; }

/* Badges */
.badge-ok  { background:#d4edda; color:#155724; padding:2px 8px; border-radius:4px; font-size:0.82rem; }
.badge-err { background:#f8d7da; color:#721c24; padding:2px 8px; border-radius:4px; font-size:0.82rem; }
.badge-warn{ background:#fff3cd; color:#856404; padding:2px 8px; border-radius:4px; font-size:0.82rem; }
</style>
""",
    unsafe_allow_html=True,
)


# ── KPI card renderer ─────────────────────────────────────────────────────────

GRADIENTS = {
    "blue":   "linear-gradient(135deg,#1e3a5f,#2980b9)",
    "green":  "linear-gradient(135deg,#1a5c38,#27ae60)",
    "red":    "linear-gradient(135deg,#6b1a1a,#e74c3c)",
    "orange": "linear-gradient(135deg,#7d4500,#f39c12)",
    "purple": "linear-gradient(135deg,#3b1a5c,#8e44ad)",
    "teal":   "linear-gradient(135deg,#0d4d4d,#1abc9c)",
}


def kpi_row(metrics: list[tuple]) -> None:
    """Render all KPI cards in one st.markdown call to avoid Streamlit
    closing-tag rendering artifacts that appear with per-column calls."""
    cards = ""
    for label, value, color, delta in metrics:
        grad = GRADIENTS.get(color, GRADIENTS["blue"])
        delta_part = f'<span class="kpi-delta">{delta}</span>' if delta else ""
        cards += (
            f'<div class="kpi-card" style="background:{grad};">'
            f'<span class="kpi-value">{value}</span>'
            f'<span class="kpi-label">{label}</span>'
            f'{delta_part}'
            f'</div>'
        )
    st.markdown(
        f'<div class="kpi-row">{cards}</div>',
        unsafe_allow_html=True,
    )


# ── File loaders ──────────────────────────────────────────────────────────────

def load_file(uploaded) -> pd.DataFrame:
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded)
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded)
    raise ValueError(f"Unsupported format: {uploaded.name}")


def available_sample_files() -> list[Path]:
    data_dir = ROOT / "data"
    return sorted(
        p for p in data_dir.iterdir()
        if p.suffix in {".csv", ".xlsx", ".xls"} and p.name != "__init__.py"
    )


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown(
        "<h2 style='color:#7fb3d3;margin-bottom:0;'>📊 Excel Automation</h2>"
        "<p style='color:#5a8aaa;font-size:0.82rem;margin-top:2px;'>Data Cleaning & Reporting Tool</p>",
        unsafe_allow_html=True,
    )
    st.divider()

    # ── Data source ──
    st.markdown("**Data Source**")
    source = st.radio("", ["Upload a file", "Use sample data"], label_visibility="collapsed")

    df_raw: pd.DataFrame | None = None
    file_label = ""

    if source == "Upload a file":
        uploaded = st.file_uploader(
            "Drop CSV or Excel here",
            type=["csv", "xlsx", "xls"],
            help="Supported: .csv · .xlsx · .xls",
        )
        if uploaded:
            try:
                df_raw = load_file(uploaded)
                file_label = uploaded.name
                logger.info("File uploaded: %s (%d rows)", file_label, len(df_raw))
            except Exception as exc:
                st.error(f"Could not read file: {exc}")
    else:
        samples = available_sample_files()
        if not samples:
            st.warning("No sample files found. Run `python data/generate_sample_data.py` first.")
        else:
            chosen = st.selectbox("Select sample", [p.name for p in samples])
            chosen_path = ROOT / "data" / chosen
            df_raw = (
                pd.read_csv(chosen_path)
                if chosen_path.suffix == ".csv"
                else pd.read_excel(chosen_path)
            )
            file_label = chosen

    if df_raw is not None:
        st.success(f"{file_label} · {len(df_raw):,} rows")
        st.divider()

        # ── Cleaning options ──
        st.markdown("**Cleaning Options**")
        opt_std_cols   = st.checkbox("Standardize column names", value=True)
        opt_trim       = st.checkbox("Trim whitespace", value=True)
        opt_dedup      = st.checkbox("Remove duplicates", value=True)
        opt_type_conv  = st.checkbox("Auto type conversion", value=True)
        missing_label  = st.selectbox(
            "Missing value strategy",
            ["Leave as-is", "Fill (mean / Unknown)", "Drop rows with any NaN", "Fill zeros"],
        )
        missing_map = {
            "Leave as-is":                "flag",
            "Fill (mean / Unknown)":      "fill_mean",
            "Drop rows with any NaN":     "drop",
            "Fill zeros":                 "fill_zero",
        }
        st.divider()

        # ── Validation config ──
        st.markdown("**Validation Config**")
        all_cols = list(df_raw.columns)
        required_cols = st.multiselect("Required columns", all_cols)
        email_cols = st.multiselect(
            "Email columns",
            all_cols,
            default=[c for c in all_cols if "email" in c.lower()],
        )
        numeric_cols_sel = st.multiselect(
            "Numeric columns to validate",
            all_cols,
            default=[
                c for c in all_cols
                if any(kw in c.lower() for kw in ["amount", "price", "qty", "quantity", "revenue", "count", "total"])
            ],
        )
        st.divider()

        process_btn = st.button("⚡  Process Data", type="primary", use_container_width=True)
    else:
        process_btn = False


# ── Main header ───────────────────────────────────────────────────────────────

st.markdown(
    """<div class="main-header">
        <h1>📊 Excel Automation &amp; Reporting Tool</h1>
        <p>Business-grade data cleaning &nbsp;·&nbsp; validation &nbsp;·&nbsp;
           analytics &nbsp;·&nbsp; Excel export</p>
    </div>""",
    unsafe_allow_html=True,
)


# ── Welcome screen ────────────────────────────────────────────────────────────

if df_raw is None:
    # Step cards
    col_a, col_b, col_c, col_d = st.columns(4)
    steps = [
        ("STEP 1", "📁 Load Data",
         "Upload your CSV or Excel file, or pick one of the included sample business datasets."),
        ("STEP 2", "🧹 Clean",
         "Remove duplicates, trim whitespace, standardize columns, and handle missing values."),
        ("STEP 3", "✅ Validate",
         "Define required columns, email fields, and numeric fields to flag bad rows automatically."),
        ("STEP 4", "📥 Export",
         "Download the cleaned CSV, flagged records, or a formatted multi-sheet Excel workbook."),
    ]
    for col, (num, title, desc) in zip([col_a, col_b, col_c, col_d], steps):
        with col:
            st.markdown(
                f"""<div class="step-card">
                    <div class="step-num">{num}</div>
                    <div class="step-title">{title}</div>
                    <div class="step-desc">{desc}</div>
                </div>""",
                unsafe_allow_html=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)

    # Feature cards
    feat_a, feat_b, feat_c = st.columns(3)
    features = [
        ("#2980b9", "🧹 Data Cleaning", [
            "Remove exact duplicate rows",
            "Handle missing values — 4 strategies",
            "Trim whitespace from all text fields",
            "Snake_case column name standardization",
            "Smart automatic type conversion",
            "Filter sparse / nearly-empty rows",
        ]),
        ("#27ae60", "✅ Data Validation", [
            "Detect missing required columns",
            "Validate email address formats (regex)",
            "Validate numeric field types",
            "Annotate each failing row with reason",
            "Completeness percentage score",
            "Duplicate row detection",
        ]),
        ("#8e44ad", "📈 Analytics & Export", [
            "Interactive Plotly charts",
            "Numeric distribution histograms",
            "Category frequency bar charts",
            "Pearson correlation heatmap",
            "Time-series line charts",
            "Formatted 4-sheet Excel workbook",
        ]),
    ]
    for col, (accent, title, items) in zip([feat_a, feat_b, feat_c], features):
        bullets = "".join(f"<li>{item}</li>" for item in items)
        with col:
            st.markdown(
                f"""<div class="feat-card" style="--accent:{accent};">
                    <div class="feat-title">{title}</div>
                    <ul>{bullets}</ul>
                </div>""",
                unsafe_allow_html=True,
            )
    st.stop()


# ── Process data ──────────────────────────────────────────────────────────────

if process_btn:
    cleaning_options = {
        "standardize_columns": opt_std_cols,
        "trim_spaces":         opt_trim,
        "remove_duplicates":   opt_dedup,
        "type_conversion":     opt_type_conv,
        "missing_strategy":    missing_map[missing_label],
    }
    # Remap column names to match the (possibly standardized) cleaned DataFrame
    def _remap(cols: list[str]) -> list[str]:
        if not opt_std_cols:
            return cols
        return [standardize_col_name(c) for c in cols]

    validation_config = {
        "required_columns": _remap(required_cols),
        "email_columns":    _remap(email_cols),
        "numeric_columns":  _remap(numeric_cols_sel),
    }

    with st.spinner("Cleaning and validating data…"):
        cleaner   = DataCleaner()
        validator = DataValidator()

        df_cleaned, df_removed = cleaner.clean(df_raw.copy(), cleaning_options)
        val_summary  = validator.generate_summary(df_cleaned, validation_config)
        df_invalid   = validator.get_invalid_rows(df_cleaned, validation_config)

        # Merge rows removed during cleaning with rows failing validation
        df_invalid_all = (
            pd.concat([df_removed, df_invalid], ignore_index=True)
            .drop_duplicates()
            .reset_index(drop=True)
        )

        analytics    = compute_analytics(df_cleaned)
        cleaning_log = cleaner.get_log()

    st.session_state.update(
        processed=True,
        df_raw=df_raw,
        df_cleaned=df_cleaned,
        df_invalid=df_invalid_all,
        val_summary=val_summary,
        analytics=analytics,
        cleaning_log=cleaning_log,
        validation_config=validation_config,
        file_label=file_label,
    )
    logger.info(
        "Processing complete — clean: %d, invalid: %d",
        len(df_cleaned), len(df_invalid_all),
    )

elif df_raw is not None and not st.session_state.get("processed"):
    st.info(
        f"File loaded: **{file_label}** — {len(df_raw):,} rows × {len(df_raw.columns)} columns.  "
        "Configure options in the sidebar then click **Process Data**."
    )
    st.dataframe(df_raw.head(20), use_container_width=True)
    st.stop()


# ── Results ───────────────────────────────────────────────────────────────────

if not st.session_state.get("processed"):
    st.stop()

# Unpack session state
_raw       = st.session_state.df_raw
_clean     = st.session_state.df_cleaned
_invalid   = st.session_state.df_invalid
_val       = st.session_state.val_summary
_analytics = st.session_state.analytics
_log       = st.session_state.cleaning_log
_vlabel    = st.session_state.file_label
_vcfg      = st.session_state.validation_config

tab_overview, tab_clean, tab_valid, tab_analytics, tab_export = st.tabs(
    ["📋 Overview", "🧹 Cleaning", "✅ Validation", "📈 Analytics", "📥 Export"]
)


# ── Tab 1: Overview ────────────────────────────────────────────────────────────

with tab_overview:
    st.markdown('<div class="section-head">Raw File Overview</div>', unsafe_allow_html=True)

    raw_info = dataframe_info(_raw)
    kpi_row([
        ("Total Rows",     format_number(raw_info["rows"]),                          "blue",   ""),
        ("Columns",        raw_info["columns"],                                       "teal",   ""),
        ("Duplicates",     format_number(int(_raw.duplicated().sum())),               "orange", ""),
        ("Missing Values", format_number(int(_raw.isnull().sum().sum())),             "red",    ""),
        ("Memory",         f"{raw_info['memory_mb']} MB",                             "purple", ""),
    ])

    st.markdown('<div class="section-head">Data Preview (first 100 rows)</div>', unsafe_allow_html=True)
    st.dataframe(_raw.head(100), use_container_width=True)

    st.markdown('<div class="section-head">Column Profile</div>', unsafe_allow_html=True)
    col_profile = pd.DataFrame({
        "Column":     _raw.columns,
        "Type":       _raw.dtypes.astype(str).values,
        "Non-Null":   _raw.count().values,
        "Null Count": _raw.isnull().sum().values,
        "Null %":     (_raw.isnull().sum() / max(len(_raw), 1) * 100).round(1).values,
        "Unique":     _raw.nunique().values,
        "Sample":     [str(_raw[c].dropna().iloc[0]) if _raw[c].notna().any() else "" for c in _raw.columns],
    })
    st.dataframe(col_profile, use_container_width=True, hide_index=True)


# ── Tab 2: Cleaning ────────────────────────────────────────────────────────────

with tab_clean:
    rows_removed = len(_raw) - len(_clean)
    dups_found   = int(_raw.duplicated().sum())

    st.markdown('<div class="section-head">Cleaning Results</div>', unsafe_allow_html=True)
    kpi_row([
        ("Original Rows",   format_number(len(_raw)),          "blue",   ""),
        ("Cleaned Rows",    format_number(len(_clean)),         "green",  ""),
        ("Rows Removed",    format_number(rows_removed),        "red",    ""),
        ("Duplicates Found",format_number(dups_found),          "orange", ""),
        ("Columns",         len(_clean.columns),                "teal",   ""),
    ])

    st.markdown('<div class="section-head">Cleaned Dataset</div>', unsafe_allow_html=True)
    st.dataframe(_clean.head(200), use_container_width=True)

    st.markdown('<div class="section-head">Processing Log</div>', unsafe_allow_html=True)
    for entry in _log:
        st.markdown(f"- {entry}")

    if not _invalid.empty:
        with st.expander(f"Rows removed during cleaning ({len(_invalid)})"):
            st.dataframe(_invalid, use_container_width=True)


# ── Tab 3: Validation ─────────────────────────────────────────────────────────

with tab_valid:
    missing_req = _val.get("missing_required_columns", [])
    inv_email   = _val.get("invalid_email_count", 0)
    inv_num     = _val.get("invalid_numeric_count", 0)
    dup_ct      = _val.get("duplicate_count", 0)
    completeness= _val.get("completeness_pct", 100)

    st.markdown('<div class="section-head">Validation Summary</div>', unsafe_allow_html=True)
    kpi_row([
        ("Missing Req. Cols", len(missing_req),       "red"    if missing_req    else "green", ""),
        ("Invalid Emails",    inv_email,               "red"    if inv_email      else "green", ""),
        ("Invalid Numeric",   inv_num,                 "red"    if inv_num        else "green", ""),
        ("Duplicates Left",   dup_ct,                  "orange" if dup_ct         else "green", ""),
        ("Completeness",      f"{completeness}%",      "green"  if completeness >= 90 else "orange", ""),
    ])

    if missing_req:
        st.error(f"Missing required columns: `{'`, `'.join(missing_req)}`")
    else:
        st.success("All required columns are present.")

    # Missing values chart
    st.markdown('<div class="section-head">Missing Values by Column</div>', unsafe_allow_html=True)
    mv = {k: v for k, v in _val.get("missing_value_counts", {}).items() if v > 0}
    if mv:
        fig = px.bar(
            x=list(mv.keys()), y=list(mv.values()),
            labels={"x": "Column", "y": "Missing Count"},
            color=list(mv.values()), color_continuous_scale="Reds",
            title="Missing Values per Column",
        )
        fig.update_layout(showlegend=False, height=340, coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.success("No missing values in the cleaned dataset.")

    # Invalid rows table
    st.markdown('<div class="section-head">Invalid Records</div>', unsafe_allow_html=True)
    if not _invalid.empty:
        st.warning(f"{len(_invalid)} row(s) flagged as invalid.")
        st.dataframe(_invalid, use_container_width=True)
    else:
        st.success("All records passed validation.")


# ── Tab 4: Analytics ──────────────────────────────────────────────────────────

with tab_analytics:
    st.markdown('<div class="section-head">Dataset Analytics</div>', unsafe_allow_html=True)

    numeric_cols  = _clean.select_dtypes(include="number").columns.tolist()
    cat_cols      = _clean.select_dtypes(include="object").columns.tolist()
    datetime_cols = _clean.select_dtypes(include=["datetime64"]).columns.tolist()

    # ── Numeric summary ──
    if numeric_cols:
        st.markdown('<div class="section-head">Numeric Summary</div>', unsafe_allow_html=True)
        st.dataframe(_clean[numeric_cols].describe().round(2), use_container_width=True)

        st.markdown('<div class="section-head">Distributions</div>', unsafe_allow_html=True)
        cols_to_plot = numeric_cols[:6]
        chart_cols = st.columns(min(len(cols_to_plot), 2))
        for i, col in enumerate(cols_to_plot):
            with chart_cols[i % 2]:
                fig = px.histogram(
                    _clean, x=col, nbins=30,
                    title=f"{col}",
                    color_discrete_sequence=["#2980b9"],
                )
                fig.update_layout(height=280, showlegend=False,
                                  margin=dict(t=35, b=20, l=20, r=20))
                st.plotly_chart(fig, use_container_width=True)

        # Correlation heatmap
        if len(numeric_cols) >= 2:
            st.markdown('<div class="section-head">Correlation Heatmap</div>', unsafe_allow_html=True)
            corr = _clean[numeric_cols].corr().round(2)
            fig = px.imshow(
                corr, text_auto=True, aspect="auto",
                color_continuous_scale="RdBu_r",
                title="Feature Correlation Matrix",
            )
            fig.update_layout(height=420)
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No numeric columns detected after cleaning.")

    # ── Categorical charts ──
    if cat_cols:
        st.markdown('<div class="section-head">Category Distribution</div>', unsafe_allow_html=True)
        # Prefer meaningful low-cardinality columns; skip ID/key/email columns
        _id_hints = {"id", "key", "email", "mail", "phone", "address", "url", "uuid"}
        useful_cats = [
            c for c in cat_cols
            if not any(h in c.lower() for h in _id_hints)
            and _clean[c].nunique() <= 50
        ] or cat_cols
        default_idx = 0
        selected_cat = st.selectbox("Select column", useful_cats, index=default_idx, key="cat_sel")
        top = _clean[selected_cat].value_counts().head(20)
        fig = px.bar(
            x=top.index, y=top.values,
            labels={"x": selected_cat, "y": "Count"},
            color=top.values, color_continuous_scale="Blues",
            title=f"Top values — {selected_cat}",
        )
        fig.update_layout(height=380, showlegend=False, coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

    # ── Time series ──
    if datetime_cols and numeric_cols:
        st.markdown('<div class="section-head">Time Series</div>', unsafe_allow_html=True)
        col_dt, col_v = st.columns(2)
        with col_dt:
            date_col = st.selectbox("Date column", datetime_cols, key="ts_date")
        with col_v:
            val_col = st.selectbox("Value column", numeric_cols, key="ts_val")

        ts = (
            _clean[[date_col, val_col]]
            .dropna()
            .groupby(date_col)[val_col]
            .sum()
            .reset_index()
        )
        fig = px.line(
            ts, x=date_col, y=val_col,
            title=f"{val_col} over time",
            color_discrete_sequence=["#2980b9"],
        )
        fig.update_layout(height=360)
        st.plotly_chart(fig, use_container_width=True)


# ── Tab 5: Export ─────────────────────────────────────────────────────────────

with tab_export:
    st.markdown('<div class="section-head">Download Reports</div>', unsafe_allow_html=True)

    exporter = ExcelExporter()

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown("**Cleaned Dataset (CSV)**")
        st.caption(f"{len(_clean):,} rows · ready to use")
        csv_bytes = _clean.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="⬇️  Download Cleaned CSV",
            data=csv_bytes,
            file_name=f"cleaned_{_vlabel.rsplit('.', 1)[0]}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with c2:
        st.markdown("**Invalid Records (CSV)**")
        st.caption(f"{len(_invalid):,} flagged row(s)")
        if not _invalid.empty:
            inv_csv = _invalid.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="⬇️  Download Invalid Records",
                data=inv_csv,
                file_name=f"invalid_{datetime.now():%Y%m%d}.csv",
                mime="text/csv",
                use_container_width=True,
            )
        else:
            st.success("No invalid records to export.")

    with c3:
        st.markdown("**Full Excel Report**")
        st.caption("4 formatted sheets · charts included")
        xlsx_bytes = exporter.export_full_report(
            cleaned_df=_clean,
            invalid_df=_invalid,
            analytics=_analytics,
            cleaning_log=_log,
        )
        st.download_button(
            label="⬇️  Download Excel Report",
            data=xlsx_bytes,
            file_name=f"report_{_vlabel.rsplit('.', 1)[0]}_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # ── Processing summary table ──
    st.divider()
    st.markdown('<div class="section-head">Processing Summary</div>', unsafe_allow_html=True)
    summary_df = pd.DataFrame(
        {
            "Metric": [
                "Source File",
                "Original Rows",
                "Cleaned Rows",
                "Rows Removed",
                "Invalid / Flagged Rows",
                "Duplicates Found",
                "Missing Values (original)",
                "Data Completeness",
                "Exported At",
            ],
            "Value": [
                _vlabel,
                f"{len(_raw):,}",
                f"{len(_clean):,}",
                f"{len(_raw) - len(_clean):,}",
                f"{len(_invalid):,}",
                f"{int(_raw.duplicated().sum()):,}",
                f"{int(_raw.isnull().sum().sum()):,}",
                f"{_val.get('completeness_pct', 100)}%",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ],
        }
    )
    st.dataframe(summary_df, use_container_width=True, hide_index=True)
