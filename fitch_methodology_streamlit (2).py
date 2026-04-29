# ============================================================
# Fitch Sovereign Methodology + Global Sovereign Data Comparator
# Streamlit App – v4 (fix XLSB parser)
# ============================================================

import io
import math
import re
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

st.set_page_config(page_title="Fitch Sovereign App", layout="wide")

INTERCEPT = 4.877

# ============================================================
# Metadata – SRM
# ============================================================

SRM_VARIABLES = {
    "structural": {
        "governance_indicator": {
            "label": "Composite governance indicator (percentile rank)",
            "coefficient": 0.079,
            "weight": 22.1,
            "help": "Simple average percentile rank of World Bank governance indicators.",
        },
        "gdp_per_capita_percentile": {
            "label": "GDP per capita percentile rank (USD, market FX)",
            "coefficient": 0.037,
            "weight": 11.7,
            "help": "Percentile rank across Fitch-rated sovereigns.",
        },
        "share_world_gdp_log": {
            "label": "Share in world GDP (natural log of % share)",
            "coefficient": 0.645,
            "weight": 14.5,
            "help": "Natural logarithm of the percentage share in world GDP.",
        },
        "years_since_default_transform": {
            "label": "Years since default/restructuring event (SRM transformed value)",
            "coefficient": -1.744,
            "weight": 4.3,
            "help": "Use 0 if no event after 1980, or input the SRM-ready transformed value.",
        },
        "money_supply_log": {
            "label": "Broad money supply (% of GDP) - natural log",
            "coefficient": 0.148,
            "weight": 1.1,
            "help": "Natural logarithm of broad money as % of GDP.",
        },
    },
    "macro": {
        "real_gdp_growth_volatility_log": {
            "label": "Real GDP growth volatility - natural log",
            "coefficient": -0.704,
            "weight": 4.5,
            "help": "Natural log of the exponentially weighted std. deviation of historical real GDP growth.",
        },
        "consumer_price_inflation": {
            "label": "Consumer price inflation (3-year centred average, %, truncated 2%-50%)",
            "coefficient": -0.068,
            "weight": 3.6,
            "help": "Public criteria truncate this variable to the 2%-50% range.",
        },
        "real_gdp_growth": {
            "label": "Real GDP growth (3-year centred average, %)",
            "coefficient": 0.054,
            "weight": 1.7,
            "help": "Three-year centred average.",
        },
    },
    "public_finances": {
        "gross_general_govt_debt": {
            "label": "Gross general government debt (% of GDP, 3-year centred average)",
            "coefficient": -0.023,
            "weight": 9.2,
            "help": "Gross general government debt, centred three-year average.",
        },
        "general_govt_interest_revenue": {
            "label": "General government interest (% of revenues, 3-year centred average)",
            "coefficient": -0.044,
            "weight": 4.6,
            "help": "General government interest expenditures as % of revenues.",
        },
        "general_govt_fiscal_balance": {
            "label": "General government fiscal balance (% of GDP, 3-year centred average)",
            "coefficient": 0.039,
            "weight": 2.1,
            "help": "General government fiscal balance, centred three-year average.",
        },
        "fc_govt_debt_share": {
            "label": "Foreign-currency government debt (% of gross government debt, 3-year centred average)",
            "coefficient": -0.008,
            "weight": 3.2,
            "help": "Foreign-currency or indexed government debt as % of gross government debt.",
        },
    },
    "external": {
        "reserve_currency_flexibility": {
            "label": "Reserve-currency flexibility (SRM transformed value)",
            "coefficient": 0.484,
            "weight": 7.1,
            "help": "Use 0 if the sovereign has no reserve-currency flexibility.",
        },
        "sovereign_net_foreign_assets": {
            "label": "Sovereign net foreign assets (% of GDP, 3-year centred average)",
            "coefficient": 0.010,
            "weight": 7.5,
            "help": "Three-year centred average of sovereign net foreign assets.",
        },
        "commodity_dependence": {
            "label": "Commodity dependence (% of current external receipts)",
            "coefficient": -0.003,
            "weight": 1.0,
            "help": "Non-manufactured merchandise exports as % of current external receipts.",
        },
        "fx_reserves_months_cxp": {
            "label": "Official FX reserves (months of CXP) - for non-reserve-currency sovereigns",
            "coefficient": 0.021,
            "weight": 1.2,
            "help": "Usually set to 0 if reserve-currency flexibility is above 0.",
        },
        "external_interest_service": {
            "label": "External interest service (% of current external receipts, 3-year centred average)",
            "coefficient": -0.004,
            "weight": 0.2,
            "help": "Three-year centred average.",
        },
        "cab_plus_net_fdi": {
            "label": "Current account balance + net inward FDI (% of GDP, 3-year centred average)",
            "coefficient": 0.004,
            "weight": 0.4,
            "help": "Three-year centred average.",
        },
    },
}

PILLAR_LABELS = {
    "structural": "I. Structural Features",
    "macro": "II. Macroeconomic Performance, Policies and Prospects",
    "public_finances": "III. Public Finances",
    "external": "IV. External Finances",
}

QO_FACTORS = {
    "structural": [
        "Political stability and capacity",
        "Financial sector risks",
        "Other structural factors",
    ],
    "macro": [
        "Macroeconomic policy credibility and flexibility",
        "GDP growth outlook",
        "Macroeconomic stability / imbalances",
    ],
    "public_finances": [
        "Fiscal financing flexibility",
        "Public debt sustainability",
        "Fiscal structure",
    ],
    "external": [
        "External financing flexibility",
        "External debt sustainability",
        "Vulnerability to shocks",
    ],
}

QO_GUIDANCE = {
    2: "Exceptionally strong features relative to SRM data and output",
    1: "Strong features relative to SRM data and output",
    0: "Average features relative to SRM data and output",
    -1: "Weak features relative to SRM data and output",
    -2: "Exceptionally weak features relative to SRM data and output",
}

RATING_SCALE_NUMERIC = {
    16: "AAA", 15: "AA+", 14: "AA", 13: "AA-",
    12: "A+", 11: "A", 10: "A-",
    9: "BBB+", 8: "BBB", 7: "BBB-",
    6: "BB+", 5: "BB", 4: "BB-",
    3: "B+", 2: "B", 1: "B-",
}

LONG_TERM_SCALE = [
    "AAA", "AA+", "AA", "AA-",
    "A+", "A", "A-",
    "BBB+", "BBB", "BBB-",
    "BB+", "BB", "BB-",
    "B+", "B", "B-",
    "CCC+", "CCC", "CCC-", "CC", "C", "RD", "D",
]

ST_MAPPING_OPTIONS = {
    "AAA": ("F1+", "F1+"), "AA+": ("F1+", "F1+"), "AA": ("F1+", "F1+"), "AA-": ("F1+", "F1+"),
    "A+": ("F1", "F1+"), "A": ("F1", "F1+"), "A-": ("F2", "F1"),
    "BBB+": ("F2", "F1"), "BBB": ("F3", "F2"), "BBB-": ("F3", "F3"),
    "BB+": ("B", "B"), "BB": ("B", "B"), "BB-": ("B", "B"),
    "B+": ("B", "B"), "B": ("B", "B"), "B-": ("B", "B"),
    "CCC+": ("C", "C"), "CCC": ("C", "C"), "CCC-": ("C", "C"),
    "CC": ("C", "C"), "C": ("C", "C"),
    "RD": ("C/RD/D", "C/RD/D"), "D": ("D", "D"),
}

VARIABLE_RULES = {
    "governance_indicator": {"step": 0.1, "soft_min": 0, "soft_max": 100},
    "gdp_per_capita_percentile": {"step": 0.1, "soft_min": 0, "soft_max": 100},
    "share_world_gdp_log": {"step": 0.1, "soft_min": -20, "soft_max": 5},
    "years_since_default_transform": {"step": 0.01, "soft_min": 0, "soft_max": 5},
    "money_supply_log": {"step": 0.1, "soft_min": -5, "soft_max": 10},
    "real_gdp_growth_volatility_log": {"step": 0.1, "soft_min": -5, "soft_max": 10},
    "consumer_price_inflation": {"step": 0.1, "soft_min": -100, "soft_max": 500},
    "real_gdp_growth": {"step": 0.1, "soft_min": -20, "soft_max": 20},
    "gross_general_govt_debt": {"step": 0.1, "soft_min": 0, "soft_max": 500},
    "general_govt_interest_revenue": {"step": 0.1, "soft_min": 0, "soft_max": 100},
    "general_govt_fiscal_balance": {"step": 0.1, "soft_min": -50, "soft_max": 50},
    "fc_govt_debt_share": {"step": 0.1, "soft_min": 0, "soft_max": 100},
    "reserve_currency_flexibility": {"step": 0.1, "soft_min": 0, "soft_max": 20},
    "sovereign_net_foreign_assets": {"step": 0.1, "soft_min": -300, "soft_max": 300},
    "commodity_dependence": {"step": 0.1, "soft_min": 0, "soft_max": 100},
    "fx_reserves_months_cxp": {"step": 0.1, "soft_min": 0, "soft_max": 60},
    "external_interest_service": {"step": 0.1, "soft_min": 0, "soft_max": 100},
    "cab_plus_net_fdi": {"step": 0.1, "soft_min": -30, "soft_max": 30},
}

# ============================================================
# SRM Helpers
# ============================================================

def clamp_qo(adjustments: dict, crisis_extension: bool = False) -> int:
    raw = int(sum(adjustments.values()))
    return raw if crisis_extension else max(-3, min(3, raw))


def score_to_lt_rating(score: float) -> str:
    rounded = int(round(score))
    if rounded >= 16:
        return "AAA"
    if rounded <= 0:
        return "CCC+"
    return RATING_SCALE_NUMERIC.get(rounded, "CCC+")


def rating_index(rating: str) -> int:
    return LONG_TERM_SCALE.index(rating) if rating in LONG_TERM_SCALE else LONG_TERM_SCALE.index("CCC+")


def apply_notches(base_rating: str, notch_adjustment: int) -> str:
    if base_rating not in LONG_TERM_SCALE:
        return base_rating
    idx = rating_index(base_rating)
    new_idx = max(0, min(len(LONG_TERM_SCALE) - 1, idx - int(notch_adjustment)))
    return LONG_TERM_SCALE[new_idx]


def map_short_term(long_rating: str, use_higher_option: bool) -> str:
    lo, hi = ST_MAPPING_OPTIONS.get(long_rating, ("C", "C"))
    return hi if use_higher_option else lo


def approx_years_since_default_transform(years_since_event=None, no_event_since_1980=True) -> float:
    if no_event_since_1980 or years_since_event is None:
        return 0.0
    years_since_event = max(0.0, float(years_since_event))
    return float(math.exp(-math.log(2) * years_since_event / 4.3))


def safe_number_input(var_key: str, label: str, default: float, help_text: str):
    rule = VARIABLE_RULES.get(var_key, {"step": 0.1})
    return st.number_input(
        label,
        value=float(default),
        step=float(rule.get("step", 0.1)),
        key=var_key,
        help=help_text,
    )


def get_clean_srm_inputs() -> dict:
    inputs = {}
    for pillar in SRM_VARIABLES.values():
        for k in pillar.keys():
            v = float(st.session_state.get(k, 0.0))
            if k == "consumer_price_inflation":
                v = min(50.0, max(2.0, v))
            inputs[k] = v
    return inputs


def compute_srm(inputs: dict) -> tuple:
    details = []
    total = INTERCEPT
    for pillar_key, vars_dict in SRM_VARIABLES.items():
        for var_key, meta in vars_dict.items():
            value = float(inputs.get(var_key, 0.0))
            contribution = value * float(meta["coefficient"])
            total += contribution
            details.append({
                "Pillar": PILLAR_LABELS[pillar_key],
                "Variable": meta["label"],
                "Value": value,
                "Coefficient": meta["coefficient"],
                "Contribution": contribution,
            })
    details.append({
        "Pillar": "Intercept",
        "Variable": "OLS intercept",
        "Value": 1.0,
        "Coefficient": INTERCEPT,
        "Contribution": INTERCEPT,
    })
    return total, details


def build_radar(srm_score: float, qo_total: int, final_score: float):
    categories = ["SRM score", "QO total", "Final model score"]
    values = [srm_score, qo_total, final_score]
    categories += categories[:1]
    values += values[:1]
    fig = go.Figure(data=[go.Scatterpolar(r=values, theta=categories, fill="toself")])
    fig.update_layout(showlegend=False, polar=dict(radialaxis=dict(visible=True)))
    return fig


# ============================================================
# State init
# ============================================================

def init_state():
    defaults = {
        "governance_indicator": 50.0,
        "gdp_per_capita_percentile": 50.0,
        "share_world_gdp_log": -4.0,
        "years_since_default_transform": 0.0,
        "money_supply_log": math.log(60.0),
        "real_gdp_growth_volatility_log": math.log(3.0),
        "consumer_price_inflation": 4.0,
        "real_gdp_growth": 3.0,
        "gross_general_govt_debt": 60.0,
        "general_govt_interest_revenue": 8.0,
        "general_govt_fiscal_balance": -3.0,
        "fc_govt_debt_share": 35.0,
        "reserve_currency_flexibility": 0.0,
        "sovereign_net_foreign_assets": -20.0,
        "commodity_dependence": 25.0,
        "fx_reserves_months_cxp": 4.0,
        "external_interest_service": 8.0,
        "cab_plus_net_fdi": -1.0,
        "qo_structural": 0,
        "qo_macro": 0,
        "qo_public_finances": 0,
        "qo_external": 0,
        "qo_crisis_extension": False,
        "lc_manual_adjust": 0,
        "fc_robust_liquidity": False,
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)


# ============================================================
# Fitch XLSB Parser
# ============================================================

def normalize_label(text: str) -> str:
    text = str(text).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def slugify(text: str) -> str:
    text = normalize_label(text).lower()
    text = text.replace("&", " and ")
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


def coerce_numeric(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.replace({
        "N/A": np.nan, "N.M.": np.nan, "NM": np.nan,
        "None": np.nan, "nan": np.nan, "": np.nan,
        "..": np.nan, "...": np.nan, "n.a": np.nan,
        "n.a.": np.nan, "-": np.nan,
    })
    return pd.to_numeric(s, errors="coerce")


@st.cache_data(show_spinner=False)
def parse_fitch_comparator(file_bytes) -> pd.DataFrame:
    """Parse the Fitch Global Sovereign Data Comparator .xlsb format."""
    from pyxlsb import open_workbook

    def _cell_str(v):
        """Convert any cell value to a clean string for matching."""
        if v is None:
            return ""
        if isinstance(v, float):
            if v == int(v):
                return str(int(v))
            return str(v)
        return str(v).strip()

    # 1. Read ALL rows from the workbook (try each sheet until we get data)
    all_rows = []
    chosen_sheet = None
    with open_workbook(io.BytesIO(file_bytes)) as wb:
        sheets = wb.sheets
        if not sheets:
            return pd.DataFrame()
        for sheet_name in sheets:
            candidate_rows = []
            with wb.get_sheet(sheet_name) as sheet:
                for row in sheet.rows():
                    candidate_rows.append([cell.v for cell in row])
            if len(candidate_rows) >= 10:
                all_rows = candidate_rows
                chosen_sheet = sheet_name
                break

    if len(all_rows) < 10:
        return pd.DataFrame()

    raw = pd.DataFrame(all_rows)
    SCAN_LIMIT = min(60, len(raw))

    # 2. Find the "Category"/"Sovereign" row OR the period/year header row
    cat_row_idx = None

    # Strategy A: look for "category" or "sovereign" as a cell value
    for i in range(SCAN_LIMIT):
        row_vals = [_cell_str(v).lower() for v in raw.iloc[i]]
        if "category" in row_vals or "sovereign" in row_vals:
            cat_row_idx = i
            break

    # Strategy B: find a row with many year-like values (2000-2035)
    if cat_row_idx is None:
        for i in range(SCAN_LIMIT):
            year_count = 0
            for v in raw.iloc[i]:
                s = _cell_str(v)
                if re.match(r"^(19|20)\d{2}$", s):
                    year_count += 1
                elif "av." in s.lower() or "average" in s.lower() or "latest" in s.lower():
                    year_count += 1
            if year_count >= 4:
                cat_row_idx = i
                break

    # Strategy C: look for numeric floats that are years (pyxlsb stores as float)
    if cat_row_idx is None:
        for i in range(SCAN_LIMIT):
            year_count = 0
            for v in raw.iloc[i]:
                if isinstance(v, (int, float)) and 2000 <= v <= 2040:
                    year_count += 1
            if year_count >= 4:
                cat_row_idx = i
                break

    if cat_row_idx is None:
        # Show diagnostic info
        diag_lines = []
        for i in range(min(30, len(raw))):
            cells_preview = [_cell_str(v)[:20] for v in raw.iloc[i][:15]]
            diag_lines.append(f"Row {i}: {cells_preview}")
        st.error("Não foi possível localizar a linha de períodos no XLSB.")
        with st.expander("Diagnóstico – primeiras 30 linhas"):
            st.code("\n".join(diag_lines))
        return pd.DataFrame()

    # 3. Extract header rows
    section_row_idx = max(0, cat_row_idx - 4)
    indicator_row_idx = max(0, cat_row_idx - 2)
    unit_row_idx = max(0, cat_row_idx - 1)
    period_row_idx = cat_row_idx

    ncols = len(raw.columns)

    def get_row_vals(idx):
        vals = []
        for c in range(ncols):
            v = raw.iloc[idx, c] if idx < len(raw) else None
            vals.append(_cell_str(v))
        return vals

    sections_raw = get_row_vals(section_row_idx)
    indicators_raw = get_row_vals(indicator_row_idx)
    units_raw = get_row_vals(unit_row_idx)
    periods_raw = get_row_vals(period_row_idx)

    # Forward-fill sections and indicators
    def forward_fill(lst):
        result = []
        current = ""
        for v in lst:
            if v and v.lower() not in ("none", "nan", ""):
                current = v
            result.append(current)
        return result

    sections = forward_fill(sections_raw)
    indicators = forward_fill(indicators_raw)

    # 4. Data rows start after cat_row
    data_start = cat_row_idx + 1

    # 5. Find first data column (after metadata columns)
    meta_end = 0
    for c in range(ncols):
        p = periods_raw[c].strip()
        if re.match(r"^(19|20)\d{2}$", p) or "av." in p.lower() or "latest" in p.lower():
            meta_end = c
            break
    if meta_end == 0:
        meta_end = 14  # fallback

    # 6. Build column metadata
    col_meta = []
    for c in range(meta_end, ncols):
        section = normalize_label(sections[c]) if c < len(sections) else ""
        indicator = normalize_label(indicators[c]) if c < len(indicators) else ""
        unit = normalize_label(units_raw[c]) if c < len(units_raw) else ""
        period = normalize_label(periods_raw[c]) if c < len(periods_raw) else ""

        if not section and not indicator:
            continue
        if not period or period.lower() in ("none", "nan", ""):
            continue

        # Parse year from period
        year_match = re.search(r"(19|20)\d{2}", period)
        year_num = int(year_match.group()) if year_match else None

        is_average = "av." in period.lower() or "average" in period.lower()
        is_forecast = year_num is not None and year_num >= 2025
        is_latest = "latest" in period.lower()

        col_meta.append({
            "col_idx": c,
            "section": section,
            "indicator": indicator,
            "unit": unit,
            "year": period,
            "year_num": year_num if year_num else 0,
            "is_average": is_average,
            "is_forecast": is_forecast,
            "is_latest": is_latest,
        })

    if not col_meta:
        st.error("Não foi possível mapear colunas de dados.")
        return pd.DataFrame()

    # 7. Parse data rows
    records = []
    for r in range(data_start, len(raw)):
        row = raw.iloc[r]

        entity_key = _cell_str(row.iloc[0])
        if not entity_key or entity_key.lower() in ("none", "nan", ""):
            continue

        # Determine entity type – col 4 typically says COUNTRY or HEADING
        entity_type_raw = _cell_str(row.iloc[4]).upper() if len(row) > 4 else ""
        if "COUNTRY" in entity_type_raw:
            entity_type = "COUNTRY"
        elif "HEADING" in entity_type_raw:
            entity_type = "HEADING"
        else:
            entity_type = "OTHER"

        # Country code (ISO) – col 5
        country_code = _cell_str(row.iloc[5]) if len(row) > 5 else ""
        # Country name – col 6
        country_name = _cell_str(row.iloc[6]) if len(row) > 6 else ""
        if not country_name or country_name.lower() in ("none", "nan"):
            country_name = entity_key

        # Rating – col 1 or col 7 (LT FC IDR)
        lt_fc_rating = ""
        for rc in [7, 1]:
            if len(row) > rc and row.iloc[rc] is not None:
                candidate = _cell_str(row.iloc[rc])
                if candidate and candidate.upper() in [r2 for r2 in LONG_TERM_SCALE] + ["NR", "WD"]:
                    lt_fc_rating = candidate.upper()
                    break
        if not lt_fc_rating:
            lt_fc_rating = _cell_str(row.iloc[1]) if len(row) > 1 else ""

        # Dev status – col 13
        dev_status = _cell_str(row.iloc[13]) if len(row) > 13 else ""
        if dev_status.lower() in ("none", "nan"):
            dev_status = ""

        # Extract data values
        for cm in col_meta:
            c = cm["col_idx"]
            if c >= len(row):
                continue
            raw_val = row.iloc[c]
            if raw_val is None:
                continue

            # Try numeric
            if isinstance(raw_val, (int, float)):
                if np.isnan(raw_val) or np.isinf(raw_val):
                    continue
                val = float(raw_val)
            else:
                s = str(raw_val).strip()
                if s.lower() in ("", "none", "nan", "n/a", "n.m.", "nm", "..", "...", "n.a", "n.a.", "-"):
                    continue
                s_clean = s.replace(",", ".")
                try:
                    val = float(s_clean)
                except ValueError:
                    continue

            records.append({
                "entity_key": entity_key,
                "entity_type": entity_type,
                "country_name": country_name,
                "country_code": country_code,
                "lt_fc_rating": lt_fc_rating,
                "dev_status": dev_status,
                "section": cm["section"],
                "indicator": cm["indicator"],
                "unit": cm["unit"],
                "year": cm["year"],
                "year_num": cm["year_num"],
                "is_average": cm["is_average"],
                "is_forecast": cm["is_forecast"],
                "value": val,
            })

    df = pd.DataFrame(records)
    if df.empty:
        return df

    # Clean up section names
    df["section"] = df["section"].replace({"": "Other"})
    df["section"] = df["section"].str.replace("&amp;", "&", regex=False)
    df["country_name"] = df["country_name"].str.replace("&amp;", "&", regex=False)
    df["indicator"] = df["indicator"].str.replace("&amp;", "&", regex=False)

    return df


# ============================================================
# Comparator -> Excel export
# ============================================================

def _sane_sheet(name, used, ml=28):
    safe = re.sub(r'[/\\?\*\[\]:]+', '_', str(name)).strip()[:ml] or 'Sheet'
    base, n = safe, 2
    while safe in used:
        safe = base[:ml-2] + '_' + str(n)
        n += 1
    return safe


def comparator_to_excel(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    used_names = set()

    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )
    header_font = Font(bold=True, size=10)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font_white = Font(bold=True, size=10, color="FFFFFF")

    sections = df["section"].unique() if "section" in df.columns else ["Data"]

    for sec in sections:
        sdf = df[df["section"] == sec].copy() if "section" in df.columns else df.copy()
        if sdf.empty:
            continue

        sheet_name = _sane_sheet(sec, used_names)
        used_names.add(sheet_name)
        ws = wb.create_sheet(title=sheet_name)

        # Pivot for the sheet
        pivot = sdf.pivot_table(
            index=["country_name", "country_code", "lt_fc_rating", "indicator"],
            columns="year",
            values="value",
            aggfunc="first"
        ).reset_index()

        # Write headers
        for c_idx, col_name in enumerate(pivot.columns, 1):
            cell = ws.cell(row=1, column=c_idx, value=str(col_name))
            cell.font = header_font_white
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.border = thin_border

        # Write data
        for r_idx, row_data in enumerate(pivot.itertuples(index=False), 2):
            for c_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = thin_border
                if isinstance(val, (int, float)):
                    cell.number_format = '#,##0.0'
                    cell.alignment = Alignment(horizontal="center")

        # Auto-width
        for col_cells in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col_cells), default=10)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 3, 30)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ============================================================
# Comparator UI functions
# ============================================================

def build_comparator_filters(df: pd.DataFrame) -> pd.DataFrame:
    st.subheader("🔍 Filtros do Comparator")

    c1, c2, c3 = st.columns(3)

    with c1:
        entity_type = st.radio(
            "Tipo de entidade",
            ["Países", "Medianas/Grupos", "Todos"],
            horizontal=True,
            key="fc_entity_type",
        )
    
    if entity_type == "Países":
        df = df[df["entity_type"] == "COUNTRY"]
    elif entity_type == "Medianas/Grupos":
        df = df[df["entity_type"] != "COUNTRY"]

    sections = sorted(df["section"].unique())
    with c2:
        sel_sections = st.multiselect(
            "Seções",
            sections,
            default=sections[:3] if len(sections) >= 3 else sections,
            key="fc_sections",
        )
    if sel_sections:
        df = df[df["section"].isin(sel_sections)]

    countries = sorted(df["country_name"].unique())
    with c3:
        sel_countries = st.multiselect(
            "Países / Grupos",
            countries,
            default=countries[:5] if len(countries) >= 5 else countries,
            key="fc_countries",
        )
    if sel_countries:
        df = df[df["country_name"].isin(sel_countries)]

    # Indicator filter
    indicators = sorted(df["indicator"].unique())
    sel_indicator = st.selectbox(
        "Indicador (para gráfico)",
        indicators,
        key="fc_indicator",
    )

    return df, sel_indicator


def render_comparator_dashboard(df: pd.DataFrame, sel_indicator: str):
    st.subheader("📊 Dashboard – Fitch Comparator")

    if df.empty:
        st.warning("Nenhum dado encontrado com os filtros selecionados.")
        return

    plot_df = df[df["indicator"] == sel_indicator].copy()
    if plot_df.empty:
        st.warning(f"Nenhum dado para o indicador '{sel_indicator}' com os filtros atuais.")
        return

    # Sort by year_num
    plot_df = plot_df.sort_values("year_num")

    # Line chart
    fig = px.line(
        plot_df,
        x="year",
        y="value",
        color="country_name",
        markers=True,
        title=f"{sel_indicator}",
        labels={"value": plot_df["unit"].iloc[0] if "unit" in plot_df.columns and len(plot_df) > 0 else "Valor",
                "year": "Período", "country_name": "País"},
    )
    fig.update_layout(
        xaxis_tickangle=-45,
        legend=dict(orientation="h", yanchor="bottom", y=-0.4),
        height=500,
    )
    st.plotly_chart(fig, use_container_width=True)

    # Stats
    with st.expander("📈 Estatísticas do indicador selecionado"):
        stats = plot_df.groupby("country_name")["value"].agg(["count", "mean", "min", "max", "std"]).reset_index()
        stats.columns = ["País", "Obs", "Média", "Mín", "Máx", "Desvio"]
        st.dataframe(stats, use_container_width=True, hide_index=True)

    # Bar chart for latest year
    latest_year = plot_df["year_num"].max()
    bar_df = plot_df[plot_df["year_num"] == latest_year]
    if not bar_df.empty:
        fig_bar = px.bar(
            bar_df.sort_values("value", ascending=True),
            x="value",
            y="country_name",
            orientation="h",
            title=f"{sel_indicator} – {bar_df['year'].iloc[0]}",
            labels={"value": bar_df["unit"].iloc[0] if len(bar_df) > 0 else "", "country_name": ""},
            color="value",
            color_continuous_scale="RdYlGn",
        )
        fig_bar.update_layout(height=max(300, len(bar_df) * 30))
        st.plotly_chart(fig_bar, use_container_width=True)


def render_comparator_table(df: pd.DataFrame):
    st.subheader("📋 Dados em tabela – Fitch Comparator")

    if df.empty:
        st.warning("Nenhum dado encontrado.")
        return

    view = st.radio(
        "Visualização",
        ["Longa (recomendada)", "Pivotada"],
        horizontal=True,
        index=0,
        key="fc_view",
    )

    display_cols = ["section", "country_name", "country_code", "lt_fc_rating",
                    "indicator", "unit", "year", "year_num", "value"]
    existing = [c for c in display_cols if c in df.columns]

    long_df = df[existing].sort_values(
        [c for c in ["section", "country_name", "indicator", "year_num"] if c in existing]
    ).copy()

    if view == "Longa (recomendada)":
        display = long_df
    else:
        idx_cols = [c for c in ["section", "country_name", "country_code", "lt_fc_rating", "indicator"] if c in df.columns]
        display = df.pivot_table(
            index=idx_cols,
            columns="year",
            values="value",
            aggfunc="first"
        ).reset_index()

    st.dataframe(display, use_container_width=True, hide_index=True)

    c1, c2 = st.columns(2)
    with c1:
        csv = display.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ CSV",
            data=csv,
            file_name="fitch_comparator.csv",
            mime="text/csv",
            key="dl_comp_csv",
        )
    with c2:
        xlsx_bytes = comparator_to_excel(long_df)
        st.download_button(
            "⬇️ Excel (.xlsx)",
            data=xlsx_bytes,
            file_name="fitch_comparator.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_comp_xlsx",
        )


# ============================================================
# Metodologia pages
# ============================================================

def render_methodology_overview():
    st.markdown("""
## Fitch Sovereign Rating Methodology

O modelo soberano da Fitch combina:

1. **SRM (Sovereign Rating Model)**: modelo quantitativo com 18 variáveis
   agrupadas em 4 pilares, gerando um score numérico.
2. **QO (Qualitative Overlay)**: ajuste qualitativo de até ±3 notches 
   (extensível em crises), aplicado pilar a pilar.
3. **Ratings finais**: LT FC IDR → LT LC IDR → ST FC IDR → ST LC IDR.

### Pilares do SRM

| Pilar | Peso aprox. |
|-------|-------------|
| I. Structural Features | ~53% |
| II. Macroeconomic Performance | ~10% |
| III. Public Finances | ~19% |
| IV. External Finances | ~17% |

### Intercepto
O intercepto OLS é **{intercept:.3f}**.

### Escala
O score do SRM mapeia para a escala de rating usando arredondamento:
- 16 = AAA, 15 = AA+, ..., 1 = B-
- Abaixo de 1 → CCC+
    """.format(intercept=INTERCEPT))


def render_methodology_pillar(pillar_key: str):
    st.subheader(PILLAR_LABELS[pillar_key])
    st.caption("Insira os valores SRM-ready para este pilar.")

    rows = []
    for var_key, meta in SRM_VARIABLES[pillar_key].items():
        value_raw = safe_number_input(
            var_key,
            meta["label"],
            st.session_state.get(var_key, 0.0),
            meta["help"],
        )
        contribution = float(value_raw) * float(meta["coefficient"])
        rows.append({
            "Variável": meta["label"],
            "Valor": value_raw,
            "Coeficiente": meta["coefficient"],
            "Peso (%)": meta["weight"],
            "Contribuição": contribution,
        })

    df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.metric("Subtotal do pilar", f"{df['Contribuição'].sum():.3f}")


def render_methodology_qo():
    st.subheader("Qualitative Overlay (QO)")
    st.caption("Ajuste qualitativo por pilar analítico. Cap típico: ±3 notches no total.")

    for pillar_key, factors in QO_FACTORS.items():
        st.markdown(f"#### {PILLAR_LABELS[pillar_key]}")
        qo_key = f"qo_{pillar_key}"

        st.selectbox(
            f"QO – {PILLAR_LABELS[pillar_key]}",
            options=list(QO_GUIDANCE.keys()),
            format_func=lambda x: f"{x:+d}  —  {QO_GUIDANCE[x]}",
            key=qo_key,
            index=2,  # default = 0
        )

        with st.expander("Fatores considerados"):
            for f in factors:
                st.write(f"- {f}")

    st.checkbox("Crisis extension (permite QO fora de ±3)", key="qo_crisis_extension")


def render_methodology_results():
    st.subheader("Resultados do Modelo")

    inputs = get_clean_srm_inputs()
    srm_score, details = compute_srm(inputs)

    # QO
    adjustments = {
        "structural": int(st.session_state.get("qo_structural", 0)),
        "macro": int(st.session_state.get("qo_macro", 0)),
        "public_finances": int(st.session_state.get("qo_public_finances", 0)),
        "external": int(st.session_state.get("qo_external", 0)),
    }
    crisis_ext = bool(st.session_state.get("qo_crisis_extension", False))
    qo_total = clamp_qo(adjustments, crisis_ext)
    final_score = srm_score + qo_total

    # Ratings
    lt_fc_idr = score_to_lt_rating(final_score)
    lc_adjust = int(st.session_state.get("lc_manual_adjust", 0))
    lt_lc_idr = apply_notches(lt_fc_idr, -lc_adjust)  # LC can be same or higher

    fc_robust = bool(st.session_state.get("fc_robust_liquidity", False))
    st_fc_idr = map_short_term(lt_fc_idr, fc_robust)
    st_lc_idr = map_short_term(lt_lc_idr, True)

    # Display
    c1, c2, c3 = st.columns(3)
    c1.metric("SRM Score", f"{srm_score:.2f}")
    c2.metric("QO Total", f"{qo_total:+d}")
    c3.metric("Final Score", f"{final_score:.2f}")

    st.divider()

    r1, r2, r3, r4 = st.columns(4)
    r1.metric("LT FC IDR", lt_fc_idr)
    r2.metric("LT LC IDR", lt_lc_idr)
    r3.metric("ST FC IDR", st_fc_idr)
    r4.metric("ST LC IDR", st_lc_idr)

    st.divider()

    # LC adjustment
    st.number_input(
        "LC notch adjustment vs FC (positivo = LC acima do FC)",
        min_value=-3, max_value=6, value=lc_adjust, step=1,
        key="lc_manual_adjust",
    )
    st.checkbox("FC: robust external liquidity (higher ST mapping)", key="fc_robust_liquidity")

    # Radar
    st.plotly_chart(build_radar(srm_score, qo_total, final_score), use_container_width=True)

    # Details table
    with st.expander("📋 Detalhes do SRM"):
        det_df = pd.DataFrame(details)
        st.dataframe(det_df, use_container_width=True, hide_index=True)
        st.metric("Score total SRM", f"{srm_score:.3f}")


# ============================================================
# MAIN
# ============================================================

def main():
    init_state()

    # Sidebar – File upload
    st.sidebar.title("Fitch Sovereign App")

    st.sidebar.caption("v4 – Metodologia + Comparator (Dashboard & Tabela)")
    st.sidebar.markdown("---")
    st.sidebar.markdown("**Carregar Fitch Comparator (.xlsb)**")
    uploaded_file = st.sidebar.file_uploader(
        "Envie o XLSB do Fitch Comparator",
        type=["xlsb"],
        key="xlsb_upload",
    )

    comparator_df = None
    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        with st.spinner("Processando XLSB..."):
            comparator_df = parse_fitch_comparator(file_bytes)
        if comparator_df is not None and not comparator_df.empty:
            st.sidebar.success(
                f"✅ {len(comparator_df):,} registros · "
                f"{comparator_df['country_name'].nunique()} entidades · "
                f"{comparator_df['indicator'].nunique()} indicadores"
            )
        else:
            st.sidebar.error("Não foi possível extrair dados do arquivo.")

    # Main tabs
    tab_met, tab_dash, tab_data = st.tabs([
        "📘 Metodologia",
        "📊 Dashboard",
        "📋 Dados",
    ])

    # ========== TAB 1: Metodologia ==========
    with tab_met:
        st.title("Fitch Sovereign Rating Methodology")

        sub_page = st.radio(
            "Seção",
            [
                "Visão geral",
                PILLAR_LABELS["structural"],
                PILLAR_LABELS["macro"],
                PILLAR_LABELS["public_finances"],
                PILLAR_LABELS["external"],
                "Qualitative Overlay (QO)",
                "Resultados",
            ],
            horizontal=False,
            key="met_subpage",
        )

        if sub_page == "Visão geral":
            render_methodology_overview()
        elif sub_page == PILLAR_LABELS["structural"]:
            render_methodology_pillar("structural")
        elif sub_page == PILLAR_LABELS["macro"]:
            render_methodology_pillar("macro")
        elif sub_page == PILLAR_LABELS["public_finances"]:
            render_methodology_pillar("public_finances")
        elif sub_page == PILLAR_LABELS["external"]:
            render_methodology_pillar("external")
        elif sub_page == "Qualitative Overlay (QO)":
            render_methodology_qo()
        elif sub_page == "Resultados":
            render_methodology_results()

    # ========== TAB 2: Dashboard ==========
    with tab_dash:
        st.title("📊 Fitch Global Sovereign Data Comparator – Dashboard")
        if comparator_df is None or comparator_df.empty:
            st.info("⬅️ Envie o arquivo XLSB na barra lateral para ativar o dashboard.")
        else:
            filtered_df, sel_indicator = build_comparator_filters(comparator_df)
            render_comparator_dashboard(filtered_df, sel_indicator)

    # ========== TAB 3: Dados ==========
    with tab_data:
        st.title("📋 Fitch Global Sovereign Data Comparator – Dados")
        if comparator_df is None or comparator_df.empty:
            st.info("⬅️ Envie o arquivo XLSB na barra lateral para visualizar os dados.")
        else:
            # Re-apply same entity type filter
            df_for_table = comparator_df.copy()
            entity_type_table = st.radio(
                "Tipo de entidade",
                ["Países", "Medianas/Grupos", "Todos"],
                horizontal=True,
                key="fc_entity_type_table",
            )
            if entity_type_table == "Países":
                df_for_table = df_for_table[df_for_table["entity_type"] == "COUNTRY"]
            elif entity_type_table == "Medianas/Grupos":
                df_for_table = df_for_table[df_for_table["entity_type"] != "COUNTRY"]

            sections_table = sorted(df_for_table["section"].unique())
            sel_sections_table = st.multiselect(
                "Seções",
                sections_table,
                default=sections_table[:3] if len(sections_table) >= 3 else sections_table,
                key="fc_sections_table",
            )
            if sel_sections_table:
                df_for_table = df_for_table[df_for_table["section"].isin(sel_sections_table)]

            countries_table = sorted(df_for_table["country_name"].unique())
            sel_countries_table = st.multiselect(
                "Países",
                countries_table,
                default=countries_table[:10] if len(countries_table) >= 10 else countries_table,
                key="fc_countries_table",
            )
            if sel_countries_table:
                df_for_table = df_for_table[df_for_table["country_name"].isin(sel_countries_table)]

            render_comparator_table(df_for_table)


if __name__ == "__main__":
    main()
