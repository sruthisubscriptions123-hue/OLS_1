"""
OLS Regression Analysis Tool  —  Streamlit App
================================================
Run:  streamlit run ols_app.py

Requirements (install once):
    pip install streamlit pandas numpy scipy statsmodels patsy plotly openpyxl python-docx
"""

import io
import json
import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st
from scipy import stats
from patsy import dmatrices, EvalEnvironment
import statsmodels.formula.api as smf
from statsmodels.stats.stattools import durbin_watson
from statsmodels.stats.diagnostic import het_breuschpagan
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ─────────────────────────────────────────────
#  PAGE CONFIG & CUSTOM CSS
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="OLS Regression Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
/* ── Base ── */
[data-testid="stAppViewContainer"] { background: #0f1117; }
[data-testid="stHeader"] { background: transparent; }
section[data-testid="stSidebar"] { display: none; }

/* ── Main padding ── */
.main .block-container { padding: 1.5rem 2rem 4rem; max-width: 1100px; }

/* ── Typography ── */
h1 { font-size: 1.6rem !important; font-weight: 600 !important; color: #e8eaf0 !important; }
h2 { font-size: 1.1rem !important; font-weight: 600 !important; color: #7fb3ff !important;
     border-left: 3px solid #4f8ef7; padding-left: 10px; margin-top: 1.4rem !important; }
h3 { font-size: 0.95rem !important; color: #8892a4 !important;
     text-transform: uppercase; letter-spacing: 1px; }
p, li { color: #c8ccd8 !important; font-size: 0.9rem; }
label { color: #8892a4 !important; font-size: 0.82rem !important; font-weight: 500 !important; }
code { background: #1e2535 !important; color: #3ecf8e !important;
       border-radius: 4px; padding: 2px 7px; font-size: 0.82rem; }

/* ── Streamlit widgets ── */
.stSelectbox > div > div,
.stMultiSelect > div > div,
.stNumberInput > div > div { background: #1e2535 !important; border: 1px solid #2a3145 !important;
    border-radius: 7px !important; color: #e8eaf0 !important; }
.stButton > button { background: transparent !important; border: 1.5px solid #3a4460 !important;
    color: #e8eaf0 !important; border-radius: 7px !important; font-size: 0.85rem !important;
    padding: 0.4rem 1.1rem !important; transition: all 0.2s; }
.stButton > button:hover { border-color: #7fb3ff !important; color: #7fb3ff !important; }
.stButton > button[kind="primary"] { background: #4f8ef7 !important; border-color: #4f8ef7 !important;
    color: #fff !important; }
.stButton > button[kind="primary"]:hover { background: #3a7de8 !important; }
.stDownloadButton > button { background: #1a3a2a !important; border: 1.5px solid #3ecf8e !important;
    color: #3ecf8e !important; border-radius: 7px !important; font-size: 0.85rem !important; }
.stDataFrame { border-radius: 8px !important; overflow: hidden; }
.stCheckbox > label > span { color: #c8ccd8 !important; }
[data-testid="stFileUploader"] { background: #161b27; border: 2px dashed #3a4460;
    border-radius: 10px; padding: 1rem; }
[data-testid="stExpander"] { background: #161b27 !important; border: 1px solid #2a3145 !important;
    border-radius: 8px !important; }

/* ── Metric cards ── */
[data-testid="stMetric"] { background: #1e2535; border: 1px solid #2a3145;
    border-radius: 8px; padding: 1rem; }
[data-testid="stMetricLabel"] { color: #8892a4 !important; font-size: 0.75rem !important;
    text-transform: uppercase; letter-spacing: 0.7px; }
[data-testid="stMetricValue"] { color: #7fb3ff !important; font-size: 1.5rem !important; }

/* ── Alert / info ── */
.stAlert { border-radius: 8px !important; }

/* ── Step pill badge ── */
.step-badge {
    display: inline-block; padding: 3px 12px; border-radius: 20px; font-size: 0.75rem;
    font-weight: 600; margin-bottom: 0.6rem; letter-spacing: 0.5px;
    background: rgba(79,142,247,0.15); color: #7fb3ff; border: 1px solid rgba(79,142,247,0.35);
}
.sig-yes { background: rgba(62,207,142,0.15); color: #3ecf8e;
           border: 1px solid rgba(62,207,142,0.4); border-radius: 4px;
           padding: 2px 8px; font-size: 0.78rem; font-weight: 600; }
.sig-no  { background: rgba(255,102,102,0.12); color: #f77;
           border: 1px solid rgba(255,102,102,0.35); border-radius: 4px;
           padding: 2px 8px; font-size: 0.78rem; font-weight: 600; }
.verdict-ok   { color: #3ecf8e; font-weight: 600; }
.verdict-warn { color: #f5a623; font-weight: 600; }

/* ── Divider ── */
hr { border-color: #2a3145 !important; margin: 1.2rem 0 !important; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  SESSION STATE INIT
# ─────────────────────────────────────────────
STEPS = ["Load Data", "Variable Types", "Model Build",
         "Coefficients", "Fit Metrics", "Diagnostics"]

defaults = {
    "current_step": 1,
    "raw_df": None,
    "confirmed_df": None,
    "col_types": {},
    "dv": None,
    "selected_ivs": [],
    "centering": {},
    "interactions": [],      # list of dicts: {vars: [a,b,c], ways: 2|3}
    "alpha": 0.05,
    "decimals": 3,
    "ci_level": 0.95,
    "patsy_formula": "",
    "model_result": None,    # dict from _run_ols()
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────
def go_step(n):
    st.session_state["current_step"] = n


def fmt(v, d=None):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    d = d if d is not None else st.session_state["decimals"]
    return f"{v:.{d}f}"


def fmt_p(v):
    if v < 0.001:
        return "<.001"
    return f"{v:.3f}"


def badge(label):
    return f'<span class="step-badge">{label}</span>'


def infer_type(series: pd.Series) -> str:
    """Infer numeric / categorical / object / datetime."""
    # Try datetime
    if pd.api.types.is_datetime64_any_dtype(series):
        return "datetime"
    # Try numeric
    try:
        converted = pd.to_numeric(series.dropna(), errors="raise")
        _ = converted  # noqa
        return "numeric"
    except Exception:
        pass
    # Categorical heuristic: ≤15 unique values OR ≤30% unique ratio
    n_uniq = series.nunique()
    n_total = series.count()
    if n_uniq <= 15 or (n_total > 0 and n_uniq / n_total <= 0.30):
        return "categorical"
    return "object"


def apply_confirmed_types(df: pd.DataFrame, col_types: dict) -> pd.DataFrame:
    df = df.copy()
    for col, ctype in col_types.items():
        if col not in df.columns:
            continue
        if ctype == "numeric":
            df[col] = pd.to_numeric(df[col], errors="coerce")
        elif ctype == "categorical":
            df[col] = df[col].astype(str).astype("category")
        elif ctype == "datetime":
            df[col] = pd.to_datetime(df[col], errors="coerce")
        else:  # object
            df[col] = df[col].astype(str)
    return df


def build_patsy_formula() -> str:
    dv = st.session_state["dv"]
    ivs = st.session_state["selected_ivs"]
    if not dv or not ivs:
        return ""
    terms = []
    for col in ivs:
        t = st.session_state["col_types"].get(col, "numeric")
        if t == "categorical":
            terms.append(f"C({col})")
        else:
            terms.append(col)
    for ix in st.session_state["interactions"]:
        valid = [v for v in ix["vars"] if v]
        if len(valid) >= 2:
            parts = []
            for v in valid:
                t = st.session_state["col_types"].get(v, "numeric")
                parts.append(f"C({v})" if t == "categorical" else v)
            terms.append(":".join(parts))
    return f"{dv} ~ {' + '.join(terms)}"


PLOTLY_DARK = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(30,37,53,0.9)",
    font=dict(color="#8892a4", family="sans-serif", size=11),
    xaxis=dict(gridcolor="#2a3145", linecolor="#2a3145", zerolinecolor="#2a3145"),
    yaxis=dict(gridcolor="#2a3145", linecolor="#2a3145", zerolinecolor="#2a3145"),
    margin=dict(t=40, b=50, l=60, r=20),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#8892a4")),
)


# ─────────────────────────────────────────────
#  OLS ENGINE
# ─────────────────────────────────────────────
def _run_ols(df: pd.DataFrame, formula: str, alpha: float, ci_level: float,
             centering: dict) -> dict:
    """Fit OLS, return a results dict with everything needed for Steps 4-6."""
    df = df.copy()

    # Apply mean-centering for numeric vars flagged by user
    for col, do_center in centering.items():
        if do_center and col in df.columns:
            df[col] = df[col] - df[col].mean()

    df = df.dropna()

    model = smf.ols(formula=formula, data=df)
    fit = model.fit()

    ci = fit.conf_int(alpha=1 - ci_level)
    params = fit.params
    pvals  = fit.pvalues
    bse    = fit.bse
    tvals  = fit.tvalues

    coef_rows = []
    for name in params.index:
        coef_rows.append({
            "term":   str(name),
            "coef":   float(params[name]),
            "se":     float(bse[name]),
            "tstat":  float(tvals[name]),
            "pval":   float(pvals[name]),
            "ci_lo":  float(ci.loc[name, 0]),
            "ci_hi":  float(ci.loc[name, 1]),
            "sig":    bool(pvals[name] < alpha),
        })

    influence   = fit.get_influence()
    resid       = fit.resid.values
    fitted_vals = fit.fittedvalues.values
    leverage    = influence.hat_matrix_diag
    cooks_d     = influence.cooks_distance[0]
    std_resid   = influence.resid_studentized_internal

    # Shapiro-Wilk (cap at 5000)
    sw_sample = resid[:5000] if len(resid) > 5000 else resid
    sw_stat, sw_p = stats.shapiro(sw_sample)

    # Durbin-Watson
    dw = float(durbin_watson(resid))

    # Breusch-Pagan
    bp_lm, bp_p, _, _ = het_breuschpagan(resid, fit.model.exog)

    # Linearity: correlation resid ~ fitted
    lin_r, lin_p = stats.pearsonr(fitted_vals, resid)

    return {
        "coef_rows":  coef_rows,
        "rsq":        float(fit.rsquared),
        "rsq_adj":    float(fit.rsquared_adj),
        "fstat":      float(fit.fvalue),
        "fpval":      float(fit.f_pvalue),
        "df_model":   int(fit.df_model),
        "df_resid":   int(fit.df_resid),
        "aic":        float(fit.aic),
        "bic":        float(fit.bic),
        "rmse":       float(np.sqrt(fit.mse_resid)),
        "n":          int(fit.nobs),
        "sw_stat":    float(sw_stat),  "sw_p":    float(sw_p),
        "dw":         dw,
        "bp_stat":    float(bp_lm),    "bp_p":    float(bp_p),
        "lin_r":      float(lin_r),    "lin_p":   float(lin_p),
        "fitted":     fitted_vals.tolist(),
        "resid":      resid.tolist(),
        "leverage":   leverage.tolist(),
        "cooks_d":    cooks_d.tolist(),
        "std_resid":  std_resid.tolist(),
        "obs_idx":    list(range(len(resid))),
        "formula":    formula,
        "dv":         formula.split("~")[0].strip(),
        "alpha":      alpha,
        "ci_level":   ci_level,
    }


# ─────────────────────────────────────────────
#  HEADER + STEPPER
# ─────────────────────────────────────────────
st.markdown("""
<div style="display:flex; align-items:center; gap:14px; margin-bottom:1.2rem;">
  <div style="width:38px;height:38px;background:#4f8ef7;border-radius:8px;
              display:flex;align-items:center;justify-content:center;
              font-size:14px;font-weight:700;color:#fff;flex-shrink:0;">OLS</div>
  <div>
    <div style="font-size:1.15rem;font-weight:600;color:#e8eaf0;">
        Regression Analysis Tool</div>
    <div style="font-size:0.78rem;color:#8892a4;">
        Ordinary Least Squares · Patsy Formula Engine · Streamlit</div>
  </div>
</div>
""", unsafe_allow_html=True)

# Progress bar
cs = st.session_state["current_step"]
st.progress(cs / 6, text=f"Step {cs} of 6 — {STEPS[cs-1]}")

# Stepper tabs
step_cols = st.columns(6)
for i, (col, name) in enumerate(zip(step_cols, STEPS)):
    step_num = i + 1
    is_current = step_num == cs
    is_done = step_num < cs
    color = "#3ecf8e" if is_done else ("#7fb3ff" if is_current else "#3a4460")
    text_color = "#3ecf8e" if is_done else ("#e8eaf0" if is_current else "#4a5568")
    prefix = "✓ " if is_done else f"{step_num}. "
    col.markdown(
        f'<div style="text-align:center;padding:6px 2px;border-bottom:2px solid {color};'
        f'color:{text_color};font-size:0.75rem;font-weight:{"600" if is_current else "400"};">'
        f'{prefix}{name}</div>',
        unsafe_allow_html=True
    )

st.markdown("<hr>", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════
#  STEP 1 — LOAD DATA
# ═══════════════════════════════════════════════════════════════════════
if cs == 1:
    st.markdown(badge("Step 1"), unsafe_allow_html=True)
    st.markdown("## Load Data")

    uploaded = st.file_uploader(
        "Upload your Excel file (.xlsx)", type=["xlsx", "xls"],
        help="Maximum 10,000 rows recommended for best performance."
    )

    if uploaded:
        try:
            xl = pd.ExcelFile(uploaded)
            sheet_names = xl.sheet_names

            c1, c2 = st.columns(2)
            with c1:
                sheet = st.selectbox("Sheet", sheet_names)
            with c2:
                header_row = st.selectbox("Header row", [0, 1, 2],
                                          format_func=lambda x: f"Row {x+1}")

            df_raw = xl.parse(sheet, header=header_row)

            # Drop fully-empty columns
            df_raw = df_raw.loc[:, df_raw.columns.notna()]
            df_raw.columns = [str(c).strip() for c in df_raw.columns]

            # Preview
            st.markdown("### Preview (first 5 rows)")
            st.dataframe(df_raw.head(), use_container_width=True, hide_index=True)

            # Dimension cards
            st.markdown("### Dataset dimensions")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Rows", f"{len(df_raw):,}")
            m2.metric("Columns", len(df_raw.columns))
            m3.metric("Missing cells", int(df_raw.isnull().sum().sum()))
            m4.metric("Sheet", sheet)

            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("✓ Confirm & proceed to variable types", type="primary"):
                st.session_state["raw_df"] = df_raw
                go_step(2)
                st.rerun()

        except Exception as e:
            st.error(f"Could not read file: {e}")
    else:
        st.info("👆 Upload an Excel file to begin.")


# ═══════════════════════════════════════════════════════════════════════
#  STEP 2 — VARIABLE TYPES
# ═══════════════════════════════════════════════════════════════════════
elif cs == 2:
    st.markdown(badge("Step 2"), unsafe_allow_html=True)
    st.markdown("## Variable Types")
    st.info(
        "Review the inferred data types below. Adjust any column before continuing — "
        "all downstream steps will use the types confirmed here."
    )

    df = st.session_state["raw_df"]
    TYPE_OPTIONS = ["numeric", "categorical", "object", "datetime"]
    TYPE_COLORS  = {
        "numeric":     "#4f8ef7",
        "categorical": "#f5a623",
        "object":      "#3ecf8e",
        "datetime":    "#c87aff",
    }

    # Build the type editor
    new_types = {}
    header_cols = st.columns([0.5, 2.5, 3, 1.5, 2])
    for h, c in zip(["#", "Variable", "Sample values", "Inferred", "Change type"], header_cols):
        c.markdown(f"<span style='color:#8892a4;font-size:0.75rem;font-weight:600;"
                   f"text-transform:uppercase;letter-spacing:.7px'>{h}</span>",
                   unsafe_allow_html=True)
    st.markdown("<hr style='margin:4px 0 10px'>", unsafe_allow_html=True)

    for i, col in enumerate(df.columns):
        inferred = infer_type(df[col])
        prev_type = st.session_state["col_types"].get(col, inferred)

        sample_vals = df[col].dropna().astype(str).head(3).tolist()
        sample_str  = ",  ".join(sample_vals) if sample_vals else "—"

        c0, c1, c2, c3, c4 = st.columns([0.5, 2.5, 3, 1.5, 2])
        c0.markdown(f"<span style='color:#4a5568;font-size:0.8rem'>{i+1}</span>",
                    unsafe_allow_html=True)
        c1.markdown(f"<code>{col}</code>", unsafe_allow_html=True)
        c2.markdown(
            f"<span style='color:#8892a4;font-size:0.78rem'>{sample_str[:60]}</span>",
            unsafe_allow_html=True
        )
        color = TYPE_COLORS.get(inferred, "#8892a4")
        c3.markdown(
            f"<span style='background:rgba(0,0,0,0.3);color:{color};"
            f"border-radius:4px;padding:2px 8px;font-size:0.75rem;"
            f"font-family:monospace'>{inferred}</span>",
            unsafe_allow_html=True
        )
        chosen = c4.selectbox(
            f"__{col}__", TYPE_OPTIONS,
            index=TYPE_OPTIONS.index(prev_type),
            key=f"type_{col}", label_visibility="collapsed"
        )
        new_types[col] = chosen

    st.markdown("<br>", unsafe_allow_html=True)
    bc1, bc2 = st.columns([1, 5])
    with bc1:
        if st.button("‹ Back"):
            go_step(1); st.rerun()
    with bc2:
        if st.button("✓ Confirm types & build model", type="primary"):
            st.session_state["col_types"] = new_types
            confirmed = apply_confirmed_types(st.session_state["raw_df"], new_types)
            st.session_state["confirmed_df"] = confirmed
            go_step(3); st.rerun()


# ═══════════════════════════════════════════════════════════════════════
#  STEP 3 — MODEL BUILD
# ═══════════════════════════════════════════════════════════════════════
elif cs == 3:
    st.markdown(badge("Step 3"), unsafe_allow_html=True)
    st.markdown("## Model Specification")

    df   = st.session_state["confirmed_df"]
    cols = list(df.columns)
    col_types = st.session_state["col_types"]
    numeric_cols = [c for c in cols if col_types.get(c) == "numeric"]

    # ── DV + Settings ─────────────────────────────────────
    st.markdown("### Dependent variable & settings")
    r1c1, r1c2, r1c3, r1c4 = st.columns(4)
    with r1c1:
        dv_options = ["— select —"] + cols
        prev_dv = st.session_state["dv"] or "— select —"
        dv_idx = dv_options.index(prev_dv) if prev_dv in dv_options else 0
        dv = st.selectbox("Dependent variable (response)", dv_options, index=dv_idx)
        dv = None if dv == "— select —" else dv
        st.session_state["dv"] = dv
    with r1c2:
        alpha = st.selectbox("Significance level (α)",
                             [0.05, 0.01, 0.10],
                             format_func=lambda x: str(x),
                             index=[0.05, 0.01, 0.10].index(st.session_state["alpha"]))
        st.session_state["alpha"] = alpha
    with r1c3:
        ci_level = st.selectbox("CI level",
                                [0.95, 0.99, 0.90],
                                format_func=lambda x: f"{int(x*100)}%",
                                index=[0.95, 0.99, 0.90].index(st.session_state["ci_level"]))
        st.session_state["ci_level"] = ci_level
    with r1c4:
        decimals = st.selectbox("Decimal places", [2, 3, 4],
                                index=[2, 3, 4].index(st.session_state["decimals"]))
        st.session_state["decimals"] = decimals

    st.markdown("---")

    # ── IVs ───────────────────────────────────────────────
    st.markdown("### Independent variables (predictors)")
    avail_ivs = [c for c in cols if c != dv]

    prev_ivs = [v for v in st.session_state["selected_ivs"] if v in avail_ivs]
    selected_ivs = st.multiselect(
        "Select predictors", avail_ivs, default=prev_ivs,
        help="All selected variables will enter the model as main effects."
    )
    st.session_state["selected_ivs"] = selected_ivs

    # Type pills for selected IVs
    if selected_ivs:
        pill_html = ""
        for iv in selected_ivs:
            t = col_types.get(iv, "numeric")
            colors = {"numeric": "#4f8ef7", "categorical": "#f5a623",
                      "object": "#3ecf8e", "datetime": "#c87aff"}
            clr = colors.get(t, "#8892a4")
            pill_html += (f'<span style="background:rgba(0,0,0,0.35);color:{clr};'
                          f'border:1px solid {clr};border-radius:4px;padding:2px 9px;'
                          f'font-size:0.75rem;font-family:monospace;margin:2px 3px 2px 0;'
                          f'display:inline-block">{iv} <span style="opacity:.6">({t})</span></span>')
        st.markdown(pill_html, unsafe_allow_html=True)

    st.markdown("---")

    # ── Mean-centering ────────────────────────────────────
    num_selected = [iv for iv in selected_ivs if col_types.get(iv) == "numeric"]
    if num_selected:
        st.markdown("### Mean-centering (numeric predictors)")
        center_cols = st.columns(min(len(num_selected), 4))
        for j, col_name in enumerate(num_selected):
            with center_cols[j % 4]:
                prev = st.session_state["centering"].get(col_name, False)
                val = st.checkbox(f"{col_name}", value=prev, key=f"center_{col_name}")
                st.session_state["centering"][col_name] = val
        st.markdown(
            "<span style='color:#8892a4;font-size:0.78rem'>"
            "Checked variables will be mean-centered before model fitting.</span>",
            unsafe_allow_html=True
        )
        st.markdown("---")

    # ── Interactions ──────────────────────────────────────
    st.markdown("### Interaction terms *(optional · max 2 terms · max 3-way)*")

    # Sync interactions list
    if "interactions" not in st.session_state:
        st.session_state["interactions"] = []

    ixs = st.session_state["interactions"]
    n_ix = len(ixs)

    # Remove any interactions beyond 2
    while len(ixs) > 2:
        ixs.pop()

    for k, ix in enumerate(ixs):
        with st.container():
            st.markdown(
                f"<div style='background:#1e2535;border:1px solid #2a3145;"
                f"border-radius:8px;padding:12px 16px;margin-bottom:8px'>",
                unsafe_allow_html=True
            )
            ix_cols = st.columns([2, 2, 2, 1, 0.5])
            with ix_cols[0]:
                va = st.selectbox(f"Var A (term {k+1})", ["—"] + selected_ivs,
                                  index=(["—"]+selected_ivs).index(ix["vars"][0])
                                  if ix["vars"][0] in ["—"]+selected_ivs else 0,
                                  key=f"ix_{k}_0")
                ixs[k]["vars"][0] = "" if va == "—" else va
            with ix_cols[1]:
                vb = st.selectbox(f"Var B (term {k+1})", ["—"] + selected_ivs,
                                  index=(["—"]+selected_ivs).index(ix["vars"][1])
                                  if ix["vars"][1] in ["—"]+selected_ivs else 0,
                                  key=f"ix_{k}_1")
                ixs[k]["vars"][1] = "" if vb == "—" else vb
            with ix_cols[2]:
                three_way = st.checkbox("3-way?", value=ix.get("ways", 2) == 3,
                                        key=f"ix_{k}_3way")
                ixs[k]["ways"] = 3 if three_way else 2
                if three_way:
                    vc = st.selectbox(f"Var C (term {k+1})", ["—"] + selected_ivs,
                                      index=(["—"]+selected_ivs).index(ix["vars"][2])
                                      if len(ix["vars"]) > 2 and ix["vars"][2] in ["—"]+selected_ivs else 0,
                                      key=f"ix_{k}_2")
                    if len(ixs[k]["vars"]) < 3:
                        ixs[k]["vars"].append("")
                    ixs[k]["vars"][2] = "" if vc == "—" else vc
                else:
                    if len(ixs[k]["vars"]) > 2:
                        ixs[k]["vars"][2] = ""
            with ix_cols[4]:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("✕", key=f"rm_ix_{k}"):
                    ixs.pop(k)
                    st.session_state["interactions"] = ixs
                    st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

    if len(ixs) < 2:
        if st.button("＋ Add interaction term"):
            ixs.append({"vars": ["", "", ""], "ways": 2})
            st.session_state["interactions"] = ixs
            st.rerun()
    else:
        st.caption("Maximum 2 interaction terms reached.")

    st.markdown("---")

    # ── Formula preview ───────────────────────────────────
    formula = build_patsy_formula()
    st.session_state["patsy_formula"] = formula
    st.markdown("### Patsy formula preview")
    if formula:
        st.code(formula, language="r")
    else:
        st.caption("Select a dependent variable and at least one predictor to preview.")

    st.markdown("<br>", unsafe_allow_html=True)
    bc1, bc2 = st.columns([1, 5])
    with bc1:
        if st.button("‹ Back"):
            go_step(2); st.rerun()
    with bc2:
        run_clicked = st.button("▶ Run OLS model", type="primary",
                                disabled=(not formula))
    if run_clicked:
        with st.spinner("Fitting OLS model…"):
            try:
                result = _run_ols(
                    df=st.session_state["confirmed_df"],
                    formula=formula,
                    alpha=alpha,
                    ci_level=ci_level,
                    centering=st.session_state["centering"],
                )
                st.session_state["model_result"] = result
                go_step(4); st.rerun()
            except Exception as e:
                st.error(f"Model error: {e}")


# ═══════════════════════════════════════════════════════════════════════
#  STEP 4 — COEFFICIENT TABLE
# ═══════════════════════════════════════════════════════════════════════
elif cs == 4:
    st.markdown(badge("Step 4"), unsafe_allow_html=True)
    st.markdown("## Model Output — Coefficients")

    r = st.session_state["model_result"]
    if r is None:
        st.warning("No model results. Please run the model first.")
        if st.button("‹ Go back to Model Build"):
            go_step(3); st.rerun()
        st.stop()

    d     = st.session_state["decimals"]
    alpha = r["alpha"]
    ci    = int(r["ci_level"] * 100)

    # ── Model spec box ─────────────────────────────────────
    with st.expander("📋 Model specification", expanded=True):
        spec_rows = [
            ("Dependent variable",     r["dv"]),
            ("Independent variables",  ", ".join(st.session_state["selected_ivs"])),
            ("Patsy formula",          r["formula"]),
            ("Significance level (α)", str(alpha)),
            ("CI level",               f"{ci}%"),
            ("Mean-centered",          ", ".join([k for k, v in st.session_state["centering"].items() if v]) or "none"),
            ("Observations used",      str(r["n"])),
        ]
        for k, v in spec_rows:
            c1, c2 = st.columns([2, 5])
            c1.markdown(f"<span style='color:#8892a4;font-size:0.82rem'>{k}</span>",
                        unsafe_allow_html=True)
            c2.markdown(f"<code style='font-size:0.8rem'>{v}</code>",
                        unsafe_allow_html=True)

    st.markdown("---")

    # ── Coefficient table ──────────────────────────────────
    ci_lo_label = f"CI {(1-r['ci_level'])/2*100:.1f}%"
    ci_hi_label = f"CI {(1-(1-r['ci_level'])/2)*100:.1f}%"

    rows = []
    for row in r["coef_rows"]:
        sig_html = ("✅ sig" if row["sig"] else "❌ n.s.")
        rows.append({
            "Term":          row["term"],
            "Coeff.":        f"{row['coef']:.{d}f}",
            "Std. Error":    f"{row['se']:.{d}f}",
            "t-stat":        f"{row['tstat']:.{d}f}",
            "p-value":       fmt_p(row["pval"]),
            ci_lo_label:     f"{row['ci_lo']:.{d}f}",
            ci_hi_label:     f"{row['ci_hi']:.{d}f}",
            "Sig.":          sig_html,
        })

    coef_df = pd.DataFrame(rows)

    # Colour rows
    def highlight_row(row):
        term = row["Term"]
        sig  = row["Sig."]
        if "sig" in sig and "n.s." not in sig:
            return ["background-color: rgba(62,207,142,0.08); border-left: 2px solid rgba(62,207,142,0.4)"] * len(row)
        elif "n.s." in sig:
            return ["background-color: rgba(255,102,102,0.07); border-left: 2px solid rgba(255,102,102,0.3)"] * len(row)
        return [""] * len(row)

    styled = (
        coef_df.style
        .apply(highlight_row, axis=1)
        .set_properties(**{"font-size": "0.82rem", "font-family": "monospace"})
        .set_table_styles([
            {"selector": "th", "props": [
                ("background-color", "#1e2535"),
                ("color", "#8892a4"),
                ("font-size", "0.75rem"),
                ("text-transform", "uppercase"),
                ("letter-spacing", "0.7px"),
            ]},
            {"selector": "td", "props": [("color", "#e8eaf0")]},
        ])
    )
    st.dataframe(styled, use_container_width=True, hide_index=True)

    st.markdown(
        "<span style='font-size:0.78rem;color:#8892a4'>"
        "🟢 green rows = p < α (significant)  &nbsp;&nbsp;"
        "🔴 red rows = p ≥ α (not significant)</span>",
        unsafe_allow_html=True
    )

    # ── F-test summary ─────────────────────────────────────
    st.markdown("---")
    st.markdown("### F-test & overall model fit")

    f_sig = r["fpval"] < alpha
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("F-statistic",  f"{r['fstat']:.3f}",
              help=f"df({r['df_model']}, {r['df_resid']})")
    m2.metric("p-value (F)",  fmt_p(r["fpval"]),
              delta="Significant" if f_sig else "Not significant",
              delta_color="normal" if f_sig else "inverse")
    m3.metric("R²",           f"{r['rsq']:.4f}")
    m4.metric("Adj. R²",      f"{r['rsq_adj']:.4f}")
    m5.metric("RMSE",         f"{r['rmse']:.{d}f}")
    m6.metric("AIC / BIC",   f"{r['aic']:.1f} / {r['bic']:.1f}")

    st.markdown("<br>", unsafe_allow_html=True)
    bc1, bc2 = st.columns([1, 5])
    with bc1:
        if st.button("‹ Modify model"):
            go_step(3); st.rerun()
    with bc2:
        if st.button("Fit metrics ›", type="primary"):
            go_step(5); st.rerun()


# ═══════════════════════════════════════════════════════════════════════
#  STEP 5 — FIT METRICS & ASSUMPTIONS
# ═══════════════════════════════════════════════════════════════════════
elif cs == 5:
    st.markdown(badge("Step 5"), unsafe_allow_html=True)
    st.markdown("## Fit Metrics & Assumption Tests")

    r     = st.session_state["model_result"]
    alpha = r["alpha"]
    d     = st.session_state["decimals"]

    if r is None:
        st.warning("Run the model first.")
        if st.button("‹ Back"):
            go_step(3); st.rerun()
        st.stop()

    # ── Metrics ────────────────────────────────────────────
    st.markdown("### Overall fit")
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("R²",       f"{r['rsq']:.4f}",
              help=f"Model explains {r['rsq']*100:.1f}% of variance in {r['dv']}.")
    m2.metric("Adj. R²",  f"{r['rsq_adj']:.4f}",
              help="Penalises for extra predictors. Use for model comparison.")
    m3.metric("RMSE",     f"{r['rmse']:.{d}f}",
              help=f"Avg. prediction error in units of {r['dv']}.")
    m4.metric("N",        f"{r['n']:,}",
              help="Rows after listwise deletion of missing values.")
    m5.metric("AIC",      f"{r['aic']:.1f}",
              help="Lower is better when comparing models on the same data.")
    m6.metric("BIC",      f"{r['bic']:.1f}",
              help="Stricter than AIC; penalises complexity more heavily.")

    st.markdown("---")
    st.markdown("### Assumption checks")

    # ── Helpers for verdict ────────────────────────────────
    def verdict_md(ok):
        return ("✅ **Supported**" if ok else "⚠️ **Caution**")

    # ──────────────────────────────────────────────────────
    # 1. Linearity
    # ──────────────────────────────────────────────────────
    with st.expander("📐 Linearity  —  Residuals vs Fitted", expanded=True):
        lin_ok = abs(r["lin_r"]) < 0.1
        st.markdown(
            f"**Pearson r (resid ~ fitted):** `{r['lin_r']:.4f}`,  "
            f"p = `{fmt_p(r['lin_p'])}`  →  {verdict_md(lin_ok)}"
        )
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=r["fitted"], y=r["resid"], mode="markers",
            marker=dict(color="#4f8ef7", opacity=0.55, size=6),
            name="Residuals"
        ))
        fig.add_hline(y=0, line_dash="dash", line_color="#3ecf8e", line_width=1.5)
        fig.update_layout(**PLOTLY_DARK,
                          xaxis_title="Fitted values", yaxis_title="Residuals",
                          title="Residuals vs Fitted")
        st.plotly_chart(fig, use_container_width=True)
        if lin_ok:
            st.success("Residuals show no systematic linear pattern with fitted values. "
                       "Linearity assumption is supported.")
        else:
            st.warning("Residuals correlate with fitted values, suggesting a possible "
                       "non-linear relationship. Consider polynomial terms or transformations.")

    # ──────────────────────────────────────────────────────
    # 2. Independence (Durbin-Watson)
    # ──────────────────────────────────────────────────────
    with st.expander("🔄 Independence  —  Durbin-Watson", expanded=True):
        dw_ok = 1.5 < r["dw"] < 2.5
        st.markdown(
            f"**Durbin-Watson statistic:** `{r['dw']:.4f}`  →  {verdict_md(dw_ok)}"
        )
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(
            x=list(range(1, len(r["resid"])+1)),
            y=r["resid"],
            mode="lines+markers",
            line=dict(color="#4f8ef7", width=1.5),
            marker=dict(color="#4f8ef7", size=4, opacity=0.6),
            name="Residuals"
        ))
        fig2.add_hline(y=0, line_dash="dash", line_color="#3ecf8e", line_width=1)
        fig2.update_layout(**PLOTLY_DARK,
                           xaxis_title="Observation index", yaxis_title="Residuals",
                           title="Residuals vs Observation order")
        st.plotly_chart(fig2, use_container_width=True)
        if dw_ok:
            st.success("DW ≈ 2 — no significant autocorrelation detected. "
                       "Residuals appear independent.")
        elif r["dw"] < 1.5:
            st.warning("DW < 1.5 — positive autocorrelation suspected. "
                       "Consider time-series methods or adding lag terms.")
        else:
            st.warning("DW > 2.5 — negative autocorrelation suspected.")

    # ──────────────────────────────────────────────────────
    # 3. Homoscedasticity (Breusch-Pagan)
    # ──────────────────────────────────────────────────────
    with st.expander("📏 Homoscedasticity  —  Breusch-Pagan", expanded=True):
        bp_ok = r["bp_p"] > alpha
        st.markdown(
            f"**LM statistic:** `{r['bp_stat']:.3f}`,  "
            f"p = `{fmt_p(r['bp_p'])}`  →  {verdict_md(bp_ok)}"
        )
        abs_std = [abs(v) for v in r["std_resid"]]
        fig3 = go.Figure()
        fig3.add_trace(go.Scatter(
            x=r["fitted"], y=abs_std, mode="markers",
            marker=dict(color="#f5a623", opacity=0.55, size=6),
            name="|Std. residual|"
        ))
        fig3.update_layout(**PLOTLY_DARK,
                           xaxis_title="Fitted values",
                           yaxis_title="|Standardised residuals|",
                           title="Scale-Location plot")
        st.plotly_chart(fig3, use_container_width=True)
        if bp_ok:
            st.success("BP test non-significant — no evidence of heteroscedasticity. "
                       "Error variance appears constant.")
        else:
            st.warning(
                f"BP test significant (p = {fmt_p(r['bp_p'])}) — heteroscedasticity detected. "
                "Consider heteroscedasticity-consistent (HC) standard errors or a WLS model."
            )

    # ──────────────────────────────────────────────────────
    # 4. Normality (Shapiro-Wilk + Q-Q)
    # ──────────────────────────────────────────────────────
    with st.expander("🔔 Normality of residuals  —  Shapiro-Wilk + Q-Q plot", expanded=True):
        sw_ok = r["sw_p"] > alpha
        st.markdown(
            f"**Shapiro-Wilk W:** `{r['sw_stat']:.4f}`,  "
            f"p = `{fmt_p(r['sw_p'])}`  →  {verdict_md(sw_ok)}"
        )
        # Q-Q plot (JS-free, pure numpy)
        sorted_resid = np.sort(r["resid"])
        n = len(sorted_resid)
        probs = (np.arange(1, n+1) - 0.5) / n
        theoretical = stats.norm.ppf(probs)
        fig4 = go.Figure()
        fig4.add_trace(go.Scatter(
            x=theoretical, y=sorted_resid, mode="markers",
            marker=dict(color="#c87aff", opacity=0.55, size=6),
            name="Sample quantiles"
        ))
        mn, mx = theoretical[0], theoretical[-1]
        fig4.add_trace(go.Scatter(
            x=[mn, mx], y=[mn, mx], mode="lines",
            line=dict(color="#3ecf8e", dash="dash", width=1.5),
            name="Normal reference"
        ))
        fig4.update_layout(**PLOTLY_DARK,
                           xaxis_title="Theoretical quantiles",
                           yaxis_title="Sample quantiles",
                           title="Normal Q-Q Plot of Residuals")
        st.plotly_chart(fig4, use_container_width=True)
        if sw_ok:
            st.success("SW test non-significant — residuals appear approximately "
                       "normally distributed. OLS t/F inference is valid.")
        else:
            st.warning(
                f"SW test significant (p = {fmt_p(r['sw_p'])}) — residuals deviate from normality. "
                "Consider robust inference, bootstrap CIs, or variable transformations."
            )

    st.markdown("<br>", unsafe_allow_html=True)
    bc1, bc2 = st.columns([1, 5])
    with bc1:
        if st.button("‹ Back"):
            go_step(4); st.rerun()
    with bc2:
        if st.button("Diagnostics ›", type="primary"):
            go_step(6); st.rerun()


# ═══════════════════════════════════════════════════════════════════════
#  STEP 6 — DIAGNOSTICS + DOWNLOAD
# ═══════════════════════════════════════════════════════════════════════
elif cs == 6:
    st.markdown(badge("Step 6"), unsafe_allow_html=True)
    st.markdown("## Model Diagnostics")

    r     = st.session_state["model_result"]
    alpha = r["alpha"]
    d     = st.session_state["decimals"]

    if r is None:
        st.warning("Run the model first.")
        if st.button("‹ Back"):
            go_step(3); st.rerun()
        st.stop()

    n           = r["n"]
    resid_thresh = 2.5
    lev_thresh   = 2 * (len(st.session_state["selected_ivs"]) + 1) / n
    cook_thresh  = 1.0

    st.info(
        f"Thresholds:  |Std. residual| > {resid_thresh} = outlier  ·  "
        f"Leverage > {lev_thresh:.4f} = high leverage  ·  Cook's D > {cook_thresh} = influential"
    )

    std_resid = np.array(r["std_resid"])
    leverage  = np.array(r["leverage"])
    cooks_d   = np.array(r["cooks_d"])
    obs       = np.array(r["obs_idx"])

    def obs_color(i):
        if cooks_d[i] > cook_thresh:          return "#f66"
        if abs(std_resid[i]) > resid_thresh:  return "#f5a623"
        if leverage[i] > lev_thresh:          return "#c87aff"
        return "#4f8ef7"

    colors = [obs_color(i) for i in obs]

    col_l, col_r = st.columns(2)

    # ── Standardised residuals ─────────────────────────────
    with col_l:
        st.markdown("#### Standardised residuals")
        fig_sr = go.Figure()
        fig_sr.add_trace(go.Scatter(
            x=obs+1, y=std_resid, mode="markers",
            marker=dict(color=colors, size=7, opacity=0.75),
            text=[f"Obs {i+1}" for i in obs], hovertemplate="%{text}<br>Std resid: %{y:.3f}<extra></extra>"
        ))
        fig_sr.add_hline(y=resid_thresh,  line_dash="dash", line_color="#f66", line_width=1.2)
        fig_sr.add_hline(y=-resid_thresh, line_dash="dash", line_color="#f66", line_width=1.2)
        fig_sr.add_hline(y=0, line_dash="dot", line_color="#3ecf8e", line_width=1)
        fig_sr.update_layout(**PLOTLY_DARK, height=300,
                             xaxis_title="Observation", yaxis_title="Std. residual")
        st.plotly_chart(fig_sr, use_container_width=True)
        st.caption("Amber = |std. resid| > 2.5 (outlier).")

    # ── Leverage ───────────────────────────────────────────
    with col_r:
        st.markdown("#### Leverage (hat values)")
        lev_colors = ["#c87aff" if leverage[i] > lev_thresh else "#4f8ef7" for i in obs]
        fig_lev = go.Figure()
        fig_lev.add_trace(go.Scatter(
            x=obs+1, y=leverage, mode="markers",
            marker=dict(color=lev_colors, size=7, opacity=0.75),
            text=[f"Obs {i+1}" for i in obs], hovertemplate="%{text}<br>Leverage: %{y:.4f}<extra></extra>"
        ))
        fig_lev.add_hline(y=lev_thresh, line_dash="dash", line_color="#c87aff", line_width=1.2)
        fig_lev.update_layout(**PLOTLY_DARK, height=300,
                              xaxis_title="Observation", yaxis_title="Leverage")
        st.plotly_chart(fig_lev, use_container_width=True)
        st.caption("Purple = high leverage (> 2(k+1)/n).")

    # ── Cook's Distance ────────────────────────────────────
    st.markdown("#### Cook's Distance")
    cook_colors = ["#f66" if cooks_d[i] > cook_thresh else "#4f8ef7" for i in obs]
    fig_cd = go.Figure()
    fig_cd.add_trace(go.Bar(
        x=obs+1, y=cooks_d, marker_color=cook_colors, opacity=0.8,
        text=[f"Obs {i+1}" for i in obs],
        hovertemplate="%{text}<br>Cook's D: %{y:.4f}<extra></extra>"
    ))
    fig_cd.add_hline(y=cook_thresh, line_dash="dash", line_color="#f66", line_width=1.5,
                     annotation_text="D = 1", annotation_position="top right")
    fig_cd.update_layout(**PLOTLY_DARK, height=280,
                         xaxis_title="Observation", yaxis_title="Cook's D")
    st.plotly_chart(fig_cd, use_container_width=True)

    # ── Influence map ──────────────────────────────────────
    st.markdown("#### Influence map  (leverage × Cook's D)")
    sizes = [max(8, min(30, abs(std_resid[i]) * 6)) for i in obs]
    fig_inf = go.Figure()
    fig_inf.add_trace(go.Scatter(
        x=leverage, y=cooks_d, mode="markers",
        marker=dict(color=colors, size=sizes, opacity=0.7,
                    line=dict(width=0.5, color="#0f1117")),
        text=[f"Obs {i+1}<br>Std resid: {std_resid[i]:.3f}<br>"
              f"Leverage: {leverage[i]:.4f}<br>Cook's D: {cooks_d[i]:.4f}"
              for i in obs],
        hovertemplate="%{text}<extra></extra>"
    ))
    fig_inf.add_vline(x=lev_thresh, line_dash="dash", line_color="#c87aff", line_width=1)
    fig_inf.add_hline(y=cook_thresh, line_dash="dash", line_color="#f66", line_width=1)
    fig_inf.update_layout(**PLOTLY_DARK, height=320,
                          xaxis_title="Leverage", yaxis_title="Cook's D",
                          title="Bubble size = |Std. residual|")
    st.plotly_chart(fig_inf, use_container_width=True)

    # ── Flagged table ──────────────────────────────────────
    st.markdown("---")
    st.markdown("#### Flagged observations")
    flagged_idx = [i for i in obs
                   if abs(std_resid[i]) > resid_thresh
                   or leverage[i] > lev_thresh
                   or cooks_d[i] > cook_thresh]

    if not flagged_idx:
        st.success("✓ No flagged observations under the current thresholds.")
    else:
        flag_rows = []
        for i in flagged_idx:
            flags = []
            if abs(std_resid[i]) > resid_thresh: flags.append("Outlier")
            if leverage[i] > lev_thresh:          flags.append("High leverage")
            if cooks_d[i] > cook_thresh:          flags.append("Influential")
            flag_rows.append({
                "Obs #":         i+1,
                "Std. Residual": f"{std_resid[i]:.3f}",
                "Leverage":      f"{leverage[i]:.4f}",
                "Cook's D":      f"{cooks_d[i]:.4f}",
                "Flags":         ", ".join(flags),
            })
        flag_df = pd.DataFrame(flag_rows)
        st.dataframe(flag_df, use_container_width=True, hide_index=True)

    # ═══════════════════════════════════════════════════════
    #  WORD REPORT DOWNLOAD
    # ═══════════════════════════════════════════════════════
    st.markdown("---")
    st.markdown("### 📥 Download consolidated report")

    def _set_cell_bg(cell, hex_color):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), hex_color)
        tcPr.append(shd)

    def generate_docx_report() -> bytes:
        r   = st.session_state["model_result"]
        d   = st.session_state["decimals"]
        alpha = r["alpha"]
        ci_pct = int(r["ci_level"] * 100)

        doc = Document()
        style = doc.styles["Normal"]
        style.font.name = "Arial"
        style.font.size = Pt(10)

        # Page margins
        for section in doc.sections:
            section.top_margin    = Cm(2)
            section.bottom_margin = Cm(2)
            section.left_margin   = Cm(2.5)
            section.right_margin  = Cm(2.5)

        # ── Title ────────────────────────────────────────
        title = doc.add_heading("OLS Regression Analysis Report", 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title.runs[0].font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

        doc.add_paragraph()

        # ── Model Specification ───────────────────────────
        doc.add_heading("Model Specification", 1)
        spec_items = [
            ("Dependent variable",     r["dv"]),
            ("Independent variables",  ", ".join(st.session_state["selected_ivs"])),
            ("Patsy formula",          r["formula"]),
            ("Significance level (α)", str(alpha)),
            ("CI level",               f"{ci_pct}%"),
            ("Mean-centered vars",     ", ".join([k for k, v in st.session_state["centering"].items() if v]) or "none"),
            ("Observations (N)",       str(r["n"])),
        ]
        tbl = doc.add_table(rows=len(spec_items), cols=2)
        tbl.style = "Table Grid"
        for i, (k, v) in enumerate(spec_items):
            tbl.rows[i].cells[0].text = k
            tbl.rows[i].cells[1].text = v
            tbl.rows[i].cells[0].paragraphs[0].runs[0].bold = True
            _set_cell_bg(tbl.rows[i].cells[0], "D9E2F3")
        doc.add_paragraph()

        # ── Step 4: Coefficients ─────────────────────────
        doc.add_heading("Step 4 — Coefficient Table", 1)
        note = doc.add_paragraph(
            f"α = {alpha} · CI level = {ci_pct}% · "
            "Green shading = significant (p < α) · Red shading = not significant"
        )
        note.runs[0].font.size = Pt(9)
        note.runs[0].font.color.rgb = RGBColor(0x55, 0x55, 0x55)

        ci_lo_lbl = f"CI {(1-r['ci_level'])/2*100:.1f}%"
        ci_hi_lbl = f"CI {(1-(1-r['ci_level'])/2)*100:.1f}%"
        hdr = ["Term", "Coeff.", "Std. Err.", "t-stat", "p-value",
               ci_lo_lbl, ci_hi_lbl, "Sig."]

        ctbl = doc.add_table(rows=1 + len(r["coef_rows"]), cols=len(hdr))
        ctbl.style = "Table Grid"

        # Header row
        for j, h in enumerate(hdr):
            cell = ctbl.rows[0].cells[j]
            cell.text = h
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].runs[0].font.size = Pt(9)
            _set_cell_bg(cell, "1F3864")
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        for i, row in enumerate(r["coef_rows"], start=1):
            vals = [
                row["term"],
                f"{row['coef']:.{d}f}",
                f"{row['se']:.{d}f}",
                f"{row['tstat']:.{d}f}",
                fmt_p(row["pval"]),
                f"{row['ci_lo']:.{d}f}",
                f"{row['ci_hi']:.{d}f}",
                "***" if row["sig"] else "n.s.",
            ]
            bg = "C6EFCE" if row["sig"] else ("FFC7CE" if row["term"] != "Intercept" else "F2F2F2")
            for j, v in enumerate(vals):
                cell = ctbl.rows[i].cells[j]
                cell.text = v
                cell.paragraphs[0].runs[0].font.size = Pt(9)
                cell.paragraphs[0].runs[0].font.name = "Courier New"
                _set_cell_bg(cell, bg)

        doc.add_paragraph()

        # F-test summary
        f_p = doc.add_paragraph()
        r1 = f_p.add_run(
            f"F({r['df_model']}, {r['df_resid']}) = {r['fstat']:.3f},  "
            f"p = {fmt_p(r['fpval'])}  |  "
            f"R² = {r['rsq']:.4f}  |  Adj. R² = {r['rsq_adj']:.4f}  |  "
            f"RMSE = {r['rmse']:.{d}f}  |  AIC = {r['aic']:.1f}  |  BIC = {r['bic']:.1f}"
        )
        r1.font.name = "Courier New"
        r1.font.size = Pt(9)
        doc.add_paragraph()

        # ── Step 5: Assumptions ───────────────────────────
        doc.add_heading("Step 5 — Fit Metrics & Assumption Tests", 1)

        def assumption_row(name, stat_str, ok):
            p = doc.add_paragraph()
            run = p.add_run(f"{name}:  {stat_str}  →  {'✓ Supported' if ok else '⚠ Caution'}")
            run.font.size = Pt(10)
            run.font.bold = True
            run.font.color.rgb = (RGBColor(0x00, 0x7A, 0x3D) if ok
                                  else RGBColor(0xC0, 0x6A, 0x00))

        assumption_row(
            "Linearity (r resid~fitted)",
            f"r = {r['lin_r']:.4f}, p = {fmt_p(r['lin_p'])}",
            abs(r["lin_r"]) < 0.1
        )
        assumption_row(
            "Independence (Durbin-Watson)",
            f"DW = {r['dw']:.4f}",
            1.5 < r["dw"] < 2.5
        )
        assumption_row(
            "Homoscedasticity (Breusch-Pagan)",
            f"LM = {r['bp_stat']:.3f}, p = {fmt_p(r['bp_p'])}",
            r["bp_p"] > alpha
        )
        assumption_row(
            "Normality (Shapiro-Wilk)",
            f"W = {r['sw_stat']:.4f}, p = {fmt_p(r['sw_p'])}",
            r["sw_p"] > alpha
        )
        doc.add_paragraph()

        # ── Step 6: Diagnostics ───────────────────────────
        doc.add_heading("Step 6 — Model Diagnostics", 1)
        doc.add_paragraph(
            f"Thresholds:  |Std. residual| > {resid_thresh} = outlier  ·  "
            f"Leverage > {lev_thresh:.4f} = high leverage  ·  "
            f"Cook's D > {cook_thresh} = influential"
        ).runs[0].font.size = Pt(9)

        flagged_idx2 = [i for i in obs
                        if abs(std_resid[i]) > resid_thresh
                        or leverage[i] > lev_thresh
                        or cooks_d[i] > cook_thresh]

        if not flagged_idx2:
            doc.add_paragraph("✓ No flagged observations under the current thresholds.")
        else:
            dhdr = ["Obs #", "Std. Residual", "Leverage", "Cook's D", "Flags"]
            dtbl = doc.add_table(rows=1 + len(flagged_idx2), cols=5)
            dtbl.style = "Table Grid"
            for j, h in enumerate(dhdr):
                cell = dtbl.rows[0].cells[j]
                cell.text = h
                cell.paragraphs[0].runs[0].bold = True
                cell.paragraphs[0].runs[0].font.size = Pt(9)
                _set_cell_bg(cell, "1F3864")
                cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

            for row_num, i in enumerate(flagged_idx2, start=1):
                flags = []
                if abs(std_resid[i]) > resid_thresh: flags.append("Outlier")
                if leverage[i] > lev_thresh:          flags.append("High leverage")
                if cooks_d[i] > cook_thresh:          flags.append("Influential")
                vals = [str(i+1), f"{std_resid[i]:.3f}", f"{leverage[i]:.4f}",
                        f"{cooks_d[i]:.4f}", ", ".join(flags)]
                bg = "FFC7CE"
                for j, v in enumerate(vals):
                    cell = dtbl.rows[row_num].cells[j]
                    cell.text = v
                    cell.paragraphs[0].runs[0].font.size = Pt(9)
                    cell.paragraphs[0].runs[0].font.name = "Courier New"
                    _set_cell_bg(cell, bg)

        doc.add_paragraph()
        end = doc.add_paragraph("— End of Report —")
        end.alignment = WD_ALIGN_PARAGRAPH.CENTER
        end.runs[0].font.color.rgb = RGBColor(0x99, 0x99, 0x99)
        end.runs[0].font.size = Pt(9)

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf.getvalue()

    col_dl, _ = st.columns([2, 4])
    with col_dl:
        docx_bytes = generate_docx_report()
        st.download_button(
            label="⬇ Download Report (.docx)",
            data=docx_bytes,
            file_name="OLS_Regression_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("‹ Back to Fit Metrics"):
        go_step(5); st.rerun()
