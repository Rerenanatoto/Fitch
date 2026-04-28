import io
import json
import math
import re
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

st.set_page_config(page_title="Fitch Sovereign Methodology + Comparator", layout="wide")

APP_DIR = Path(__file__).resolve().parent
ASSETS_DIR = APP_DIR / "assets"
INTERCEPT = 4.877

# ============================================================
# Metadata
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
    16: "AAA",
    15: "AA+",
    14: "AA",
    13: "AA-",
    12: "A+",
    11: "A",
    10: "A-",
    9: "BBB+",
    8: "BBB",
    7: "BBB-",
    6: "BB+",
    5: "BB",
    4: "BB-",
    3: "B+",
    2: "B",
    1: "B-",
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
    "AAA": ("F1+", "F1+"),
    "AA+": ("F1+", "F1+"),
    "AA": ("F1+", "F1+"),
    "AA-": ("F1+", "F1+"),
    "A+": ("F1", "F1+"),
    "A": ("F1", "F1+"),
    "A-": ("F2", "F1"),
    "BBB+": ("F2", "F1"),
    "BBB": ("F3", "F2"),
    "BBB-": ("F3", "F3"),
    "BB+": ("B", "B"),
    "BB": ("B", "B"),
    "BB-": ("B", "B"),
    "B+": ("B", "B"),
    "B": ("B", "B"),
    "B-": ("B", "B"),
    "CCC+": ("C", "C"),
    "CCC": ("C", "C"),
    "CCC-": ("C", "C"),
    "CC": ("C", "C"),
    "C": ("C", "C"),
    "RD": ("C/RD/D", "C/RD/D"),
    "D": ("D", "D"),
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
# Helpers
# ============================================================

def show_image(path: Path, caption: str | None = None):
    try:
        st.image(str(path), caption=caption, use_container_width=True)
    except TypeError:
        try:
            st.image(str(path), caption=caption, use_column_width=True)
        except TypeError:
            st.image(str(path), caption=caption)


def clamp_qo(adjustments: dict[str, int], crisis_extension: bool = False) -> int:
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


def approx_years_since_default_transform(years_since_event: float | None, no_event_since_1980: bool) -> float:
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


def get_clean_srm_inputs() -> dict[str, float]:
    """
    Returns the SRM inputs cleaned for calculation.
    Important: do NOT write back into st.session_state here.
    """
    inputs = {}
    for pillar in SRM_VARIABLES.values():
        for k in pillar.keys():
            v = float(st.session_state.get(k, 0.0))
            if k == "consumer_price_inflation":
                v = min(50.0, max(2.0, v))
            inputs[k] = v
    return inputs


def validate_inputs(inputs: dict[str, float]) -> list[str]:
    warnings = []

    for var_key, value in inputs.items():
        rule = VARIABLE_RULES.get(var_key)
        if not rule:
            continue
        soft_min = rule.get("soft_min")
        soft_max = rule.get("soft_max")

        if soft_min is not None and value < soft_min:
            warnings.append(f"**{var_key}** está abaixo da faixa esperada ({value:.2f} < {soft_min}).")
        if soft_max is not None and value > soft_max:
            warnings.append(f"**{var_key}** está acima da faixa esperada ({value:.2f} > {soft_max}).")

    if inputs.get("reserve_currency_flexibility", 0) > 0 and inputs.get("fx_reserves_months_cxp", 0) != 0:
        warnings.append("`reserve_currency_flexibility > 0` normalmente implica `fx_reserves_months_cxp = 0` no SRM.")

    if inputs.get("governance_indicator", 0) > 100 or inputs.get("gdp_per_capita_percentile", 0) > 100:
        warnings.append("Percentis normalmente devem ficar entre 0 e 100.")

    if inputs.get("fc_govt_debt_share", 0) > 100:
        warnings.append("`fc_govt_debt_share` normalmente não deve ultrapassar 100%.")

    if inputs.get("commodity_dependence", 0) > 100:
        warnings.append("`commodity_dependence` normalmente não deve ultrapassar 100%.")

    return warnings


def compute_srm(inputs: dict[str, float]) -> tuple[float, list[dict]]:
    details = []
    total = INTERCEPT
    for pillar_key, vars_dict in SRM_VARIABLES.items():
        for var_key, meta in vars_dict.items():
            value = float(inputs.get(var_key, 0.0))
            contribution = value * float(meta["coefficient"])
            total += contribution
            details.append(
                {
                    "Pillar": PILLAR_LABELS[pillar_key],
                    "Variable": meta["label"],
                    "Value": value,
                    "Coefficient": meta["coefficient"],
                    "Contribution": contribution,
                }
            )
    details.append(
        {
            "Pillar": "Intercept",
            "Variable": "OLS intercept",
            "Value": 1.0,
            "Coefficient": INTERCEPT,
            "Contribution": INTERCEPT,
        }
    )
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
# State
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
# Pages
# ============================================================

def page_overview():
    st.title("Fitch Sovereign Methodology")
    st.caption("Blindada: SRM + QO + Results com validações leves de entrada.")

    col1, col2 = st.columns([1.1, 0.9])

    with col1:
        st.markdown(
            """
### Incluído
- SRM starting point
- QO pelos 4 pilares
- LT FC / LT LC / ST FC / ST LC
- Validações leves de entrada

### Notas
- Usa inputs “SRM-ready” quando a variável pública é transformada
- Saídas em **CCC+ ou abaixo** devem ser tratadas como orientação de modelo
            """
        )

    with col2:
        img = ASSETS_DIR / "fitch_summary.png"
        if img.exists():
            show_image(img, "Optional local summary image")
        else:
            st.info("No local image found in assets/.")


def page_srm_inputs(pillar: str):
    st.title(PILLAR_LABELS[pillar])
    st.caption("Enter SRM-ready values for this pillar.")

    if pillar == "structural":
        with st.expander("Helpers", expanded=False):
            share_pct = st.number_input(
                "Share in world GDP (% share, raw)",
                min_value=0.000001,
                value=0.5,
                step=0.1,
                key="helper_share_world_pct",
            )
            st.write(f"log(% share) ≈ **{math.log(max(share_pct, 1e-9)):.4f}**")

            broad_money_pct = st.number_input(
                "Broad money (% of GDP, raw)",
                min_value=0.0001,
                value=60.0,
                step=1.0,
                key="helper_broad_money",
            )
            st.write(f"log(broad money % GDP) ≈ **{math.log(max(broad_money_pct, 1e-9)):.4f}**")

            no_event = st.checkbox(
                "No default/restructuring event after 1980",
                value=True,
                key="helper_no_event",
            )
            yrs = st.number_input(
                "Years since last default/restructuring event",
                min_value=0.0,
                value=0.0,
                step=1.0,
                key="helper_years_since_event",
            )
            approx = approx_years_since_default_transform(yrs, no_event)
            st.write(f"Approx transformed value ≈ **{approx:.4f}**")

            if st.button("Use approximate transformed value", key="use_years_transform"):
                # Safe because this button is above the actual widget creation for years_since_default_transform
                st.session_state["years_since_default_transform"] = approx

    if pillar == "macro":
        with st.expander("Helper: volatility log", expanded=False):
            vol_raw = st.number_input(
                "Real GDP growth volatility (raw std. dev.)",
                min_value=0.0001,
                value=3.0,
                step=0.1,
                key="helper_real_gdp_vol",
            )
            st.write(f"log(volatility) ≈ **{math.log(max(vol_raw, 1e-9)):.4f}**")

    if pillar == "external":
        with st.expander("External notes", expanded=False):
            st.markdown(
                """
- `reserve_currency_flexibility > 0` normalmente implica `fx_reserves_months_cxp = 0` no SRM.
- `commodity_dependence` e percentuais semelhantes normalmente não devem ultrapassar 100.
                """
            )

    rows = []

    for var_key, meta in SRM_VARIABLES[pillar].items():
        value_raw = safe_number_input(
            var_key,
            meta["label"],
            st.session_state.get(var_key, 0.0),
            meta["help"],
        )

        value_used = float(value_raw)

        if var_key == "consumer_price_inflation":
            clipped = min(50.0, max(2.0, value_used))
            if clipped != value_used:
                st.warning(
                    f"{meta['label']} foi ajustada automaticamente para {clipped:.2f}%, "
                    "conforme a regra pública 2%-50%."
                )
            value_used = clipped

        rule = VARIABLE_RULES.get(var_key, {})
        soft_min = rule.get("soft_min")
        soft_max = rule.get("soft_max")

        if soft_min is not None and value_used < soft_min:
            st.warning(
                f"{meta['label']} está abaixo da faixa esperada "
                f"({value_used:.2f} < {soft_min})."
            )
        if soft_max is not None and value_used > soft_max:
            st.warning(
                f"{meta['label']} está acima da faixa esperada "
                f"({value_used:.2f} > {soft_max})."
            )

        rows.append(
            {
                "Variable": meta["label"],
                "Coefficient": meta["coefficient"],
                "SRM weight (%)": meta["weight"],
                "Input value": value_used,
                "Contribution": value_used * float(meta["coefficient"]),
            }
        )

    df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.metric("Pillar subtotal", f"{df['Contribution'].sum():.3f}")


def page_qo():
    st.title("Qualitative Overlay (QO)")
    st.caption("QO by analytical pillar. Typical overall cap: +3/-3.")

    cols = st.columns(4)
    pillar_keys = ["structural", "macro", "public_finances", "external"]

    for col, pillar_key in zip(cols, pillar_keys):
        with col:
            st.subheader(PILLAR_LABELS[pillar_key].split(". ", 1)[-1])
            for item in QO_FACTORS[pillar_key]:
                st.markdown(f"- {item}")
            choice = st.selectbox(
                f"QO notch – {PILLAR_LABELS[pillar_key]}",
                options=[2, 1, 0, -1, -2],
                index=[2, 1, 0, -1, -2].index(st.session_state.get(f"qo_{pillar_key}", 0)),
                key=f"qo_{pillar_key}",
                format_func=lambda x: f"{x:+d} — {QO_GUIDANCE[x]}",
            )
            st.text_area("Rationale", key=f"qo_note_{pillar_key}", height=90)

    crisis = st.checkbox("Allow overall QO to exceed +3/-3 in exceptional circumstances", key="qo_crisis_extension")

    raw = {k: int(st.session_state.get(f"qo_{k}", 0)) for k in pillar_keys}
    qo_total = clamp_qo(raw, crisis_extension=crisis)

    c1, c2 = st.columns(2)
    c1.metric("Raw QO total", f"{sum(raw.values()):+d}")
    c2.metric("Applied QO total", f"{qo_total:+d}")

    if not crisis and qo_total != sum(raw.values()):
        st.info("Raw QO total capped to +3/-3.")


def page_results():
    st.title("Results")

    srm_inputs = get_clean_srm_inputs()
    srm_score, details = compute_srm(srm_inputs)

    qo_raw = {
        "structural": int(st.session_state.get("qo_structural", 0)),
        "macro": int(st.session_state.get("qo_macro", 0)),
        "public_finances": int(st.session_state.get("qo_public_finances", 0)),
        "external": int(st.session_state.get("qo_external", 0)),
    }
    qo_total = clamp_qo(qo_raw, crisis_extension=bool(st.session_state.get("qo_crisis_extension", False)))
    final_model_score = srm_score + qo_total

    implied_lt_fc = score_to_lt_rating(final_model_score)
    srm_lt_fc = score_to_lt_rating(srm_score)

    c1, c2, c3 = st.columns(3)
    c1.metric("SRM score", f"{srm_score:.3f}")
    c2.metric("Applied QO total", f"{qo_total:+d}")
    c3.metric("Final model score", f"{final_model_score:.3f}")

    c1, c2 = st.columns(2)
    c1.metric("LT FC IDR from SRM only", srm_lt_fc)
    c2.metric("LT FC IDR after QO", implied_lt_fc)

    if implied_lt_fc == "CCC+" or final_model_score <= 0:
        st.warning("Output at CCC+ or below should be treated as model guidance only.")

    warnings = validate_inputs(srm_inputs)
    if warnings:
        with st.expander("Warnings / consistency checks", expanded=True):
            for w in warnings:
                st.warning(w)

    lc_adjust = st.selectbox(
        "Manual LC rating notch difference versus FC",
        options=[-2, -1, 0, 1, 2],
        index=[-2, -1, 0, 1, 2].index(int(st.session_state.get("lc_manual_adjust", 0))),
        key="lc_manual_adjust",
        format_func=lambda x: f"{x:+d} notch(es)",
    )
    lt_lc = apply_notches(implied_lt_fc, lc_adjust)

    robust_liquidity = st.checkbox("Use higher short-term FC mapping option", key="fc_robust_liquidity")
    st_fc = map_short_term(implied_lt_fc, robust_liquidity)
    st_lc = map_short_term(lt_lc, True)

    r1, r2, r3, r4 = st.columns(4)
    r1.metric("LT FC IDR", implied_lt_fc)
    r2.metric("LT LC IDR", lt_lc)
    r3.metric("ST FC IDR", st_fc)
    r4.metric("ST LC IDR", st_lc)

    df = pd.DataFrame(details).sort_values(["Pillar", "Contribution"], ascending=[True, False])
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.plotly_chart(build_radar(srm_score, qo_total, final_model_score), use_container_width=True)

    payload = {
        "srm_inputs": srm_inputs,
        "srm_score": srm_score,
        "lt_fc_srm_only": srm_lt_fc,
        "qo_raw": qo_raw,
        "qo_applied": qo_total,
        "final_model_score": final_model_score,
        "lt_fc_idr": implied_lt_fc,
        "lt_lc_idr": lt_lc,
        "st_fc_idr": st_fc,
        "st_lc_idr": st_lc,
        "qo_notes": {k: st.session_state.get(f"qo_note_{k}", "") for k in qo_raw.keys()},
        "warnings": warnings,
    }
    st.download_button(
        "Download JSON snapshot",
        data=json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"),
        file_name="fitch_srm_qo_snapshot.json",
        mime="application/json",
    )


# ============================================================
# Fitch Global Sovereign Data Comparator – Parser
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
    })
    return pd.to_numeric(s, errors="coerce")


def parse_fitch_comparator(file_bytes=None) -> pd.DataFrame:
    """Parse the Fitch Global Sovereign Data Comparator .xlsx/.xlsb format.
    
    Structure (single 'Data' sheet):
      Row 5  (idx=5): section headers (forward-fill) – e.g. 'DOMESTIC ECONOMY', 'GOVERNMENT FINANCE'
      Row 7  (idx=7): indicator names (forward-fill)
      Row 8  (idx=8): units / detail line
      Row 9  (idx=9): year labels (2024, 2025, 2026, '19-23 av.', etc.)
      Row 10+ (idx>=10): data rows
        col 0: internal Fitch key
        col 1: rating category
        col 2: region code
        col 3: sub-code
        col 4: 'COUNTRY' or 'HEADING'
        col 5: ISO code
        col 6: country name
        col 7: (empty or extra)
        col 8: LT FC IDR
        col 9+: dates, flags, then data columns
    """
    empty = pd.DataFrame(columns=[
        "section", "indicator", "indicator_detail", "country_name",
        "country_code", "lt_fc_rating", "year", "year_num",
        "is_forecast", "value"
    ])

    if file_bytes is not None:
        source = io.BytesIO(file_bytes)
    else:
        data_dir = Path(__file__).resolve().parent / "data"
        app_dir = Path(__file__).resolve().parent
        candidates = list(data_dir.glob("*.xlsx")) + list(data_dir.glob("*.xlsb")) + \
                     list(app_dir.glob("*.xlsx")) + list(app_dir.glob("*.xlsb"))
        candidates = [p for p in candidates if not p.name.startswith("~$")]
        if not candidates:
            return empty
        source = candidates[0]

    # Try to read Data sheet
    try:
        engine = "pyxlsb" if str(source).endswith(".xlsb") or (hasattr(source, 'name') and str(source.name).endswith(".xlsb")) else "openpyxl"
    except Exception:
        engine = "openpyxl"

    try:
        xls = pd.ExcelFile(source, engine=engine)
    except Exception:
        try:
            xls = pd.ExcelFile(source, engine="openpyxl")
        except Exception:
            return empty

    sheet_name = None
    for s in xls.sheet_names:
        if s.strip().lower() == "data":
            sheet_name = s
            break
    if sheet_name is None:
        return empty

    raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    if raw.shape[0] < 11 or raw.shape[1] < 15:
        return empty

    # --- Build column map from header rows ---
    # Row 5: sections (forward-fill)
    section_row = raw.iloc[5].copy()
    section_row = section_row.where(section_row.notna()).ffill()

    # Row 7: indicator names (forward-fill)
    indicator_row = raw.iloc[7].copy()
    indicator_row = indicator_row.where(indicator_row.notna()).ffill()

    # Row 8: detail/units
    detail_row = raw.iloc[8].copy()

    # Row 9: year labels
    year_row = raw.iloc[9].copy()

    # Find which row has 'COUNTRY' in col 4 to confirm data start
    col4 = raw.iloc[:, 4].astype(str).str.strip().str.upper() if raw.shape[1] > 4 else pd.Series()

    # Determine data column range: skip metadata columns (0-based)
    # Find the first column where section_row has a non-null, non-metadata value
    SKIP_SECTIONS = {"RATINGS AND METADATA", "Macro", "nan", ""}
    data_col_start = None
    for ci in range(raw.shape[1]):
        sec = str(section_row.iloc[ci]).strip() if pd.notna(section_row.iloc[ci]) else ""
        if sec and sec not in SKIP_SECTIONS:
            data_col_start = ci
            break
    if data_col_start is None:
        return empty

    # Build column metadata
    col_meta = []
    for ci in range(data_col_start, raw.shape[1]):
        sec = str(section_row.iloc[ci]).strip() if pd.notna(section_row.iloc[ci]) else ""
        ind = str(indicator_row.iloc[ci]).strip() if pd.notna(indicator_row.iloc[ci]) else ""
        det = str(detail_row.iloc[ci]).strip() if pd.notna(detail_row.iloc[ci]) else ""
        yr = str(year_row.iloc[ci]).strip() if pd.notna(year_row.iloc[ci]) else ""
        if sec in SKIP_SECTIONS or not ind or ind == "nan" or not yr or yr == "nan":
            continue
        # Combine indicator + detail for uniqueness
        full_indicator = ind
        if det and det != "nan" and det != ind:
            full_indicator = f"{ind} {det}"
        col_meta.append({
            "col_idx": ci,
            "section": sec,
            "indicator": normalize_label(full_indicator),
            "indicator_detail": det if det != "nan" else "",
            "year": yr,
        })

    if not col_meta:
        return empty

    col_map_df = pd.DataFrame(col_meta)

    # --- Extract data rows ---
    # Determine country_type column, iso column, name column, rating column
    # Based on observed structure: col4=type, col5=ISO, col6=name
    # LT FC IDR is in col 8
    # But we need to find the exact columns dynamically
    
    # Find the row index where 'Category' or 'COUNTRY' first appears in col 4
    type_col = 4
    iso_col = 5
    name_col = 6
    
    # Find rating column: look for "LT FC IDR" in row 7 or row 8
    rating_col = None
    for ci in range(raw.shape[1]):
        val7 = str(raw.iloc[7, ci]).strip() if ci < raw.shape[1] else ""
        val8 = str(raw.iloc[8, ci]).strip() if ci < raw.shape[1] else ""
        if "LT FC IDR" in val7 or "LT FC IDR" in val8:
            rating_col = ci
            break
    if rating_col is None:
        # Fallback: try col 8
        rating_col = 8 if raw.shape[1] > 8 else None

    # Extract data starting from row 10
    data_start_row = 10
    data_rows = raw.iloc[data_start_row:].copy()
    
    # Filter to COUNTRY rows only (skip HEADING rows)
    if data_rows.shape[1] > type_col:
        mask = data_rows.iloc[:, type_col].astype(str).str.strip().str.upper() == "COUNTRY"
        data_rows = data_rows[mask].copy()
    
    if data_rows.empty:
        return empty

    # Build records
    records = []
    data_col_indices = col_map_df["col_idx"].tolist()
    
    for _, row in data_rows.iterrows():
        country_name = str(row.iloc[name_col]).strip() if pd.notna(row.iloc[name_col]) else ""
        iso_code = str(row.iloc[iso_col]).strip() if pd.notna(row.iloc[iso_col]) else ""
        lt_fc = str(row.iloc[rating_col]).strip() if rating_col is not None and pd.notna(row.iloc[rating_col]) else ""
        
        if not country_name or country_name == "nan":
            continue
            
        for _, cm in col_map_df.iterrows():
            ci = cm["col_idx"]
            val = row.iloc[ci] if ci < len(row) else None
            records.append({
                "section": cm["section"],
                "indicator": cm["indicator"],
                "indicator_detail": cm["indicator_detail"],
                "country_name": country_name,
                "country_code": iso_code,
                "lt_fc_rating": lt_fc,
                "year": cm["year"],
                "value_raw": val,
            })

    if not records:
        return empty

    df = pd.DataFrame(records)
    df["value"] = coerce_numeric(df["value_raw"])
    df["year_num"] = pd.to_numeric(
        df["year"].astype(str).str.extract(r"(\d{4})")[0], errors="coerce"
    )
    df["is_forecast"] = df["year"].astype(str).str.contains(r"[ef]$", case=False, na=False)
    
    return df[[
        "section", "indicator", "indicator_detail", "country_name",
        "country_code", "lt_fc_rating", "year", "year_num",
        "is_forecast", "value"
    ]].dropna(subset=["indicator", "year"])


# ============================================================
# Comparator → Excel export
# ============================================================

def _sane_sheet(name, used, ml=28):
    safe = re.sub(r'[/\\?*\[\]:]+', '_', str(name)).strip()[:ml] or 'Sheet'
    base, n = safe, 2
    while safe in used:
        safe = base[:ml-2] + '_' + str(n); n += 1
    return safe


def comparator_to_excel(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    ws_data = wb.create_sheet('Dados')
    hdr_fill = PatternFill('solid', fgColor='1F4E79')
    hdr_font = Font(bold=True, color='FFFFFF')
    thin = Side(style='thin', color='BFBFBF')
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)
    alt = PatternFill('solid', fgColor='D6E4F0')

    cols_export = [c for c in
        ['section', 'country_name', 'country_code', 'lt_fc_rating',
         'indicator', 'year', 'year_num', 'value', 'is_forecast']
        if c in df.columns]
    for ci, col in enumerate(cols_export, 1):
        c = ws_data.cell(1, ci, col)
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = Alignment(horizontal='center', wrap_text=True)
        c.border = brd
    for ri, row in enumerate(df[cols_export].itertuples(index=False), 2):
        fill = alt if ri % 2 == 0 else PatternFill()
        for ci, val in enumerate(row, 1):
            import math as _math
            v = val if not (isinstance(val, float) and _math.isnan(val)) else None
            c = ws_data.cell(ri, ci, v)
            c.fill = fill; c.border = brd
            c.alignment = Alignment(horizontal='center')
    ws_data.freeze_panes = 'A2'

    # Charts
    ws_charts = wb.create_sheet('Graficos')
    used = list(wb.sheetnames)
    chart_row = 1
    indicators = sorted(df['indicator'].dropna().unique()) if 'indicator' in df.columns else []
    for ind in indicators:
        df_i = (df[df['indicator'] == ind]
                .dropna(subset=['year_num', 'value'])
                .sort_values(['country_name', 'year_num']))
        if df_i.empty:
            continue
        countries = sorted(df_i['country_name'].unique().tolist())
        years = sorted(df_i['year_num'].unique().tolist())
        n_yr = len(years); n_ct = len(countries)
        if n_ct > 30 or n_yr == 0:
            continue

        aux_name = _sane_sheet(ind, used); used.append(aux_name)
        ws = wb.create_sheet(aux_name)
        ws.cell(1, 1, 'Ano')
        for ci, ct in enumerate(countries, 2):
            ws.cell(1, ci, ct)
        for ri, yr in enumerate(years, 2):
            ws.cell(ri, 1, str(int(yr)))
            for ci, ct in enumerate(countries, 2):
                m = df_i[(df_i['country_name'] == ct) & (df_i['year_num'] == yr)]['value']
                ws.cell(ri, ci, round(float(m.mean()), 4) if not m.empty else None)

        lc = LineChart()
        lc.title = str(ind)[:45]
        lc.style = 10; lc.width = 22; lc.height = 14; lc.smooth = False

        for ci in range(2, n_ct + 2):
            lc.add_data(Reference(ws, min_col=ci, min_row=1, max_row=n_yr + 1),
                        titles_from_data=True)
        lc.set_categories(Reference(ws, min_col=1, min_row=2, max_row=n_yr + 1))
        for s in lc.series:
            s.marker.symbol = 'circle'
            s.marker.size = 8
        lc.x_axis.title = 'Ano'
        lc.y_axis.title = 'Valor'
        ws_charts.add_chart(lc, 'A' + str(chart_row))
        chart_row += 25

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ============================================================
# Comparator UI
# ============================================================

def build_comparator_filters(df: pd.DataFrame) -> pd.DataFrame:
    st.subheader("Filtros do Comparator")

    all_sections = sorted(df["section"].dropna().unique().tolist())
    sel_sections = st.multiselect("Seção", options=all_sections, default=all_sections, key="fc_sections")
    df1 = df[df["section"].isin(sel_sections)] if sel_sections else df.copy()

    all_ratings = sorted(df1["lt_fc_rating"].dropna().unique().tolist())
    sel_ratings = st.multiselect("LT FC IDR", options=all_ratings, default=[], key="fc_ratings",
                                  help="Deixe vazio para todos.")
    df2 = df1[df1["lt_fc_rating"].isin(sel_ratings)] if sel_ratings else df1.copy()

    all_countries = sorted(df2["country_name"].dropna().unique().tolist())
    sel_countries = st.multiselect("País", options=all_countries, default=[], key="fc_countries",
                                    help="Deixe vazio para todos.")

    all_ind = sorted(df2["indicator"].dropna().unique().tolist())
    sel_ind = st.multiselect("Indicadores", options=all_ind, default=[], key="fc_indicators",
                              help="Deixe vazio para todos.")

    valid_years = df2["year_num"].dropna()
    yr_min = int(valid_years.min()) if not valid_years.empty else 2019
    yr_max = int(valid_years.max()) if not valid_years.empty else 2028
    sel_yrs = st.slider("Faixa de anos", min_value=yr_min, max_value=yr_max,
                         value=(yr_min, yr_max), key="fc_years")

    filtered = df2.copy()
    if sel_countries:
        filtered = filtered[filtered["country_name"].isin(sel_countries)]
    if sel_ind:
        filtered = filtered[filtered["indicator"].isin(sel_ind)]
    filtered = filtered[filtered["year_num"].between(sel_yrs[0], sel_yrs[1], inclusive="both")]
    return filtered


def render_comparator_dashboard(df: pd.DataFrame):
    st.subheader("Dashboards – Fitch Comparator")
    if df.empty:
        st.warning("Nenhum dado encontrado com os filtros selecionados.")
        return

    c1, c2, c3 = st.columns(3)
    c1.metric("Países", df["country_name"].nunique())
    c2.metric("Indicadores", df["indicator"].nunique())
    c3.metric("Observações", f"{len(df):,}".replace(",", "."))

    sections = sorted(df["section"].dropna().unique().tolist())
    if not sections:
        st.info("Nenhuma seção disponível.")
        return

    sel_section = st.selectbox("Seção para gráficos", sections, index=0, key="fc_chart_section")
    plot_df = df[df["section"] == sel_section].dropna(subset=["year_num", "value"])
    if plot_df.empty:
        st.info("Sem dados numéricos para esta seção.")
        return

    indicators = sorted(plot_df["indicator"].dropna().unique().tolist())
    st.markdown(f"### {sel_section} — **{len(indicators)}** indicadores")
    show_legend = plot_df["country_name"].nunique() <= 12

    for i in range(0, len(indicators), 2):
        row_inds = indicators[i:i+2]
        row_cols = st.columns(2)
        for col, ind in zip(row_cols, row_inds):
            with col:
                ind_df = plot_df[plot_df["indicator"] == ind].sort_values(["country_name", "year_num"])
                if ind_df.empty:
                    st.caption(f"Sem dados: {ind}")
                    continue
                fig = px.line(ind_df, x="year_num", y="value", color="country_name",
                              markers=True, hover_data=["lt_fc_rating", "year", "country_code"],
                              title=ind)
                fig.update_layout(height=340, margin=dict(l=10, r=10, t=50, b=10),
                                  legend_title_text="País", showlegend=show_legend)
                fig.update_xaxes(title="Ano")
                fig.update_yaxes(title="Valor")
                try:
                    st.plotly_chart(fig, use_container_width=True)
                except TypeError:
                    st.plotly_chart(fig)

    st.markdown("---")
    st.download_button(
        "⬇️ Baixar Excel (.xlsx com gráficos)",
        data=comparator_to_excel(df),
        file_name="fitch_comparator_dashboard.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_comp_dash",
    )


def render_comparator_table(df: pd.DataFrame):
    st.subheader("Dados em tabela – Fitch Comparator")
    if df.empty:
        st.warning("Nenhum dado encontrado.")
        return
    view = st.radio("Visualização", ["Longa (recomendada)", "Pivotada"], horizontal=True,
                     index=0, key="fc_view")
    long_df = df.sort_values(["section", "country_name", "indicator", "year_num"]).copy()
    if view == "Longa (recomendada)":
        display = long_df
    else:
        display = df.pivot_table(
            index=["section", "country_name", "country_code", "lt_fc_rating", "indicator"],
            columns="year", values="value", aggfunc="first"
        ).reset_index()
    try:
        st.dataframe(display, use_container_width=True, hide_index=True)
    except TypeError:
        st.dataframe(display)
    csv = display.to_csv(index=False).encode("utf-8-sig")
    _c1, _c2 = st.columns(2)
    with _c1:
        st.download_button("⬇️ CSV", data=csv, file_name="fitch_comparator.csv",
                            mime="text/csv", key="dl_comp_csv")
    with _c2:
        st.download_button("⬇️ Excel (.xlsx)", data=comparator_to_excel(long_df),
                            file_name="fitch_comparator.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_comp_xlsx")


# ============================================================
# Main
# ============================================================

def main():
    init_state()

    st.sidebar.title("Fitch Sovereign App")

    # --- Comparator file ---
    st.sidebar.markdown("### Comparator data")
    uploaded = st.sidebar.file_uploader("Upload .xlsx/.xlsb", type=["xlsx", "xlsb"], key="comp_upload")
    comp_bytes = uploaded.getvalue() if uploaded is not None else None
    comp_df = parse_fitch_comparator(comp_bytes)
    comp_filtered = None

    if not comp_df.empty:
        st.sidebar.success(f"{comp_df['country_name'].nunique()} países, "
                           f"{comp_df['indicator'].nunique()} indicadores carregados")

    # --- Navigation ---
    pages = [
        "Overview",
        PILLAR_LABELS["structural"],
        PILLAR_LABELS["macro"],
        PILLAR_LABELS["public_finances"],
        PILLAR_LABELS["external"],
        "Qualitative Overlay (QO)",
        "Results",
    ]
    if not comp_df.empty:
        pages += ["Comparator – Dashboards", "Comparator – Dados"]

    page = st.sidebar.radio("Navigate", pages)

    # SRM pages
    if page == "Overview":
        page_overview()
    elif page == PILLAR_LABELS["structural"]:
        page_srm_inputs("structural")
    elif page == PILLAR_LABELS["macro"]:
        page_srm_inputs("macro")
    elif page == PILLAR_LABELS["public_finances"]:
        page_srm_inputs("public_finances")
    elif page == PILLAR_LABELS["external"]:
        page_srm_inputs("external")
    elif page == "Qualitative Overlay (QO)":
        page_qo()
    elif page == "Results":
        page_results()

    # Comparator pages
    elif page == "Comparator – Dashboards":
        with st.expander("Filtros do Comparator", expanded=True):
            comp_filtered = build_comparator_filters(comp_df)
        render_comparator_dashboard(comp_filtered if comp_filtered is not None else comp_df)

    elif page == "Comparator – Dados":
        with st.expander("Filtros do Comparator", expanded=True):
            comp_filtered = build_comparator_filters(comp_df)
        render_comparator_table(comp_filtered if comp_filtered is not None else comp_df)


if __name__ == "__main__":
    main()
