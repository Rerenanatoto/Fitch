
import json
import math
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Fitch Sovereign Methodology", layout="wide")

APP_DIR = Path(__file__).resolve().parent
ASSETS_DIR = APP_DIR / "assets"

# ============================================================
# Fitch Sovereign Methodology - interactive tool
# ------------------------------------------------------------
# IMPORTANT IMPLEMENTATION NOTE
# This app reproduces the public Fitch Sovereign Rating Criteria logic
# as far as it is operational from the published criteria. Where Fitch's
# public criteria describe a transformed SRM variable but do not provide
# a fully reproducible live dataset or exact operational constant, the app
# asks the user to enter the SRM-ready value directly or uses a clearly
# marked approximation helper.
# ============================================================

INTERCEPT = 4.877

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
            "help": "Natural logarithm of percentage share in world GDP in USD.",
        },
        "years_since_default_transform": {
            "label": "Years since default/restructuring event (SRM transformed value)",
            "coefficient": -1.744,
            "weight": 4.3,
            "help": "Use 0 if no event after 1980. Otherwise enter the transformed value used in the SRM or use the helper approximation below.",
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
            "help": "Natural log of the exponentially weighted standard deviation of historical annual real GDP growth.",
        },
        "consumer_price_inflation": {
            "label": "Consumer price inflation (three-year centred average, %, truncated 2%-50%)",
            "coefficient": -0.068,
            "weight": 3.6,
            "help": "Use the three-year centred average; app truncates to 2%-50%.",
        },
        "real_gdp_growth": {
            "label": "Real GDP growth (three-year centred average, %)",
            "coefficient": 0.054,
            "weight": 1.7,
            "help": "Use the three-year centred average.",
        },
    },
    "public_finances": {
        "gross_general_govt_debt": {
            "label": "Gross general government debt (% of GDP, three-year centred average)",
            "coefficient": -0.023,
            "weight": 9.2,
            "help": "Gross general government debt, centred three-year average.",
        },
        "general_govt_interest_revenue": {
            "label": "General government interest (% of revenues, three-year centred average)",
            "coefficient": -0.044,
            "weight": 4.6,
            "help": "General government interest expenditures as % of revenues, centred three-year average.",
        },
        "general_govt_fiscal_balance": {
            "label": "General government fiscal balance (% of GDP, three-year centred average)",
            "coefficient": 0.039,
            "weight": 2.1,
            "help": "General government fiscal balance, centred three-year average.",
        },
        "fc_govt_debt_share": {
            "label": "Foreign-currency government debt (% of gross government debt, three-year centred average)",
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
            "help": "Public criteria describe this as a transformed variable based on the share of the currency in global FX reserves; enter the SRM-ready value. Use 0 if the sovereign has no reserve-currency flexibility.",
        },
        "sovereign_net_foreign_assets": {
            "label": "Sovereign net foreign assets (% of GDP, three-year centred average)",
            "coefficient": 0.010,
            "weight": 7.5,
            "help": "Three-year centred average of sovereign net foreign assets % of GDP.",
        },
        "commodity_dependence": {
            "label": "Commodity dependence (% of current external receipts)",
            "coefficient": -0.003,
            "weight": 1.0,
            "help": "Non-manufactured merchandise exports as % of current external receipts.",
        },
        "fx_reserves_months_cxp": {
            "label": "Official FX reserves (months of CXP) - only for non-reserve-currency sovereigns",
            "coefficient": 0.021,
            "weight": 1.2,
            "help": "For reserve-currency sovereigns, this variable is typically set to zero in the SRM.",
        },
        "external_interest_service": {
            "label": "External interest service (% of current external receipts, three-year centred average)",
            "coefficient": -0.004,
            "weight": 0.2,
            "help": "Three-year centred average.",
        },
        "cab_plus_net_fdi": {
            "label": "Current account balance + net inward FDI (% of GDP, three-year centred average)",
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
    "CCC+", "CCC", "CCC-", "CC", "C", "RD", "D"
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

RR_OPTIONS = ["RR1", "RR2", "RR3", "RR4", "RR5", "RR6"]
INSTRUMENT_RATING_TABLE = {
    "RR1": {"B+": "BB-", "B": "B+", "B-": "B", "CCC+": "B-", "CCC": "CCC+", "CCC-": "CCC", "CC": "CC", "C/RD": "CC"},
    "RR2": {"B+": "BB-", "B": "B+", "B-": "B", "CCC+": "B-", "CCC": "CCC+", "CCC-": "CCC", "CC": "CCC-", "C/RD": "CCC-"},
    "RR3": {"B+": "BB-", "B": "B+", "B-": "B", "CCC+": "B-", "CCC": "CCC+", "CCC-": "CCC", "CC": "CCC-", "C/RD": "CC"},
    "RR4": {"B+": "B+", "B": "B", "B-": "B-", "CCC+": "CCC+", "CCC": "CCC", "CCC-": "CCC-", "CC": "CC", "C/RD": "C"},
    "RR5": {"B+": "B", "B": "B-", "B-": "CCC+", "CCC+": "CCC", "CCC": "CCC-", "CCC-": "CC", "CC": "C", "C/RD": "C"},
    "RR6": {"B+": "B-", "B": "CCC+", "B-": "CCC", "CCC+": "CCC-", "CCC": "CC", "CCC-": "C", "CC": "C", "C/RD": "C"},
}

QO_GUIDANCE = {
    2: "Exceptionally strong features relative to SRM data and output",
    1: "Strong features relative to SRM data and output",
    0: "Average features relative to SRM data and output",
    -1: "Weak features relative to SRM data and output",
    -2: "Exceptionally weak features relative to SRM data and output",
}


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
    if crisis_extension:
        return raw
    return max(-3, min(3, raw))


def score_to_lt_rating(score: float) -> str:
    rounded = int(round(score))
    if rounded >= 16:
        return "AAA"
    if rounded <= 0:
        return "CCC+"
    return RATING_SCALE_NUMERIC.get(rounded, "CCC+")


def rating_index(rating: str) -> int:
    if rating not in LONG_TERM_SCALE:
        return LONG_TERM_SCALE.index("CCC+")
    return LONG_TERM_SCALE.index(rating)


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
    """
    Approximation helper only.
    Public criteria say the variable is non-linear, equals 1 in the year of the event,
    is zero if there has been no event after 1980, and that the effect halves in about 4.3 years.
    This helper implements a simple exponential decay with half-life 4.3 years.
    """
    if no_event_since_1980 or years_since_event is None:
        return 0.0
    years_since_event = max(0.0, float(years_since_event))
    return float(math.exp(-math.log(2) * years_since_event / 4.3))


def compute_srm(inputs: dict[str, float]) -> tuple[float, list[dict]]:
    details = []
    total = INTERCEPT
    for pillar_key, vars_dict in SRM_VARIABLES.items():
        pillar_sum = 0.0
        for var_key, meta in vars_dict.items():
            value = float(inputs.get(var_key, 0.0))
            contribution = value * float(meta["coefficient"])
            pillar_sum += contribution
            details.append(
                {
                    "Pillar": PILLAR_LABELS[pillar_key],
                    "Variable": meta["label"],
                    "Value": value,
                    "Coefficient": meta["coefficient"],
                    "Contribution": contribution,
                }
            )
        total += pillar_sum
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


def debt_instrument_rating(base_idr: str, rr: str) -> str:
    key = "C/RD" if base_idr in {"C", "RD", "D"} else base_idr
    return INSTRUMENT_RATING_TABLE.get(rr, {}).get(key, base_idr)


def build_radar(srm_score: float, qo_total: int, final_score: float):
    categories = ["SRM score", "QO total", "Final model score"]
    values = [srm_score, qo_total, final_score]
    categories += categories[:1]
    values += values[:1]
    fig = go.Figure(data=[go.Scatterpolar(r=values, theta=categories, fill="toself")])
    fig.update_layout(showlegend=False, polar=dict(radialaxis=dict(visible=True)))
    return fig


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
        "rr_choice": "RR4",
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)


def page_overview():
    st.title("Fitch Sovereign Rating Criteria – Interactive Tool")
    st.caption("Public-criteria Streamlit app based on Fitch's 15 September 2025 Sovereign Rating Criteria.")

    col1, col2 = st.columns([1.1, 0.9])
    with col1:
        st.markdown(
            """
### What this app does
- Uses the **Sovereign Rating Model (SRM)** as the starting point.
- Applies a **Qualitative Overlay (QO)** by the four Fitch analytical pillars.
- Produces a **model-implied LT FC IDR starting point**, then lets you derive LC/ST ratings.
- Includes a **debt instrument / recovery** page for sovereign bonds at lower rating levels.

### Fitch logic embedded
- **Four analytical pillars**: Structural Features, Macro Performance/Policies/Prospects, Public Finances, and External Finances.
- **18 SRM variables** plus the **OLS intercept**.
- **QO** with a potential **+2/-2** per pillar and a typical overall cap of **+3/-3**.
- LT FC IDR mapping from the SRM score to **AAA ... CCC+**.
            """
        )
    with col2:
        img = ASSETS_DIR / "fitch_summary.png"
        if img.exists():
            show_image(img, "Optional local summary image")
        else:
            st.info("No local image found in assets/. The app works without images.")

    with st.expander("Implementation notes", expanded=True):
        st.markdown(
            """
- Some public Fitch SRM variables are already **transformed** (for example, percentile ranks or logarithms).
- For variables where the public criteria do **not** provide a fully operational constant or live comparator dataset, this app asks for the **SRM-ready value** directly.
- The **years since default/restructuring event** helper is an **approximation** based on the criteria's description that the variable starts at 1 in the event year and halves roughly every 4.3 years.
- When the proposed LT FC IDR is **CCC+ or below**, Fitch says ratings are directly based on its definitions rather than explained through SRM + QO. This app therefore flags those cases as **model guidance only**.
            """
        )


def page_srm_inputs(pillar: str):
    st.title(PILLAR_LABELS[pillar])
    st.caption("Enter SRM-ready values for the variables in this analytical pillar.")

    if pillar == "structural":
        with st.expander("Helper: transform raw values into SRM-style inputs", expanded=False):
            share_pct = st.number_input("Share in world GDP (% share, raw)", min_value=0.000001, value=0.5, step=0.1, key="helper_share_world_pct")
            st.write(f"Natural log(% share) ≈ **{math.log(max(share_pct, 1e-9)):.4f}**")
            broad_money_pct = st.number_input("Broad money (% of GDP, raw)", min_value=0.0001, value=60.0, step=1.0, key="helper_broad_money")
            st.write(f"Natural log(broad money % GDP) ≈ **{math.log(max(broad_money_pct, 1e-9)):.4f}**")
            no_event = st.checkbox("No default/restructuring event after 1980", value=True, key="helper_no_event")
            yrs = st.number_input("Years since last default/restructuring event", min_value=0.0, value=0.0, step=1.0, key="helper_years_since_event")
            approx = approx_years_since_default_transform(yrs, no_event)
            st.write(f"Approximate transformed value ≈ **{approx:.4f}**")

    if pillar == "macro":
        with st.expander("Helper: transform raw volatility into SRM-style log input", expanded=False):
            vol_raw = st.number_input("Real GDP growth volatility (raw standard deviation)", min_value=0.0001, value=3.0, step=0.1, key="helper_real_gdp_vol")
            st.write(f"Natural log(volatility) ≈ **{math.log(max(vol_raw, 1e-9)):.4f}**")

    if pillar == "external":
        with st.expander("Helper notes for external inputs", expanded=False):
            st.markdown(
                """
- **Reserve-currency flexibility** is an SRM transformed variable in the public criteria. Use **0** if the sovereign has no reserve-currency flexibility.
- The **FX reserves (months of CXP)** variable is used only for sovereigns **without reserve-currency flexibility**.
                """
            )

    entries = []
    for var_key, meta in SRM_VARIABLES[pillar].items():
        default = float(st.session_state.get(var_key, 0.0))
        min_v = -1000.0 if var_key in {"general_govt_fiscal_balance", "sovereign_net_foreign_assets", "cab_plus_net_fdi"} else 0.0
        value = st.number_input(meta["label"], value=default, step=0.1, key=var_key, help=meta["help"], min_value=min_v)
        if var_key == "consumer_price_inflation":
            truncated = min(50.0, max(2.0, float(value)))
            st.caption(f"Inflation used by SRM after truncation: **{truncated:.2f}%**")
            value = truncated
            st.session_state[var_key] = truncated
        entries.append(
            {
                "Variable": meta["label"],
                "Coefficient": meta["coefficient"],
                "SRM weight (%)": meta["weight"],
                "Input value": float(value),
                "Contribution": float(value) * float(meta["coefficient"]),
            }
        )

    df = pd.DataFrame(entries)
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.metric("Pillar subtotal", f"{df['Contribution'].sum():.3f}")


def page_qo():
    st.title("Qualitative Overlay (QO)")
    st.caption("Apply Fitch QO notching by analytical pillar. Typical overall cap: +3/-3, except certain crisis circumstances.")

    cols = st.columns(4)
    pillar_keys = ["structural", "macro", "public_finances", "external"]
    for col, pillar_key in zip(cols, pillar_keys):
        with col:
            st.subheader(PILLAR_LABELS[pillar_key].split('. ', 1)[-1])
            st.write("**Illustrative factors**")
            for item in QO_FACTORS[pillar_key]:
                st.markdown(f"- {item}")
            choice = st.selectbox(
                f"QO notch adjustment – {PILLAR_LABELS[pillar_key]}",
                options=[2, 1, 0, -1, -2],
                index=[2, 1, 0, -1, -2].index(st.session_state.get(f"qo_{pillar_key}", 0)),
                key=f"qo_{pillar_key}",
                format_func=lambda x: f"{x:+d} — {QO_GUIDANCE[x]}",
            )
            st.caption(QO_GUIDANCE[choice])
            st.text_area("Rationale", key=f"qo_note_{pillar_key}", height=120)

    st.markdown("---")
    crisis = st.checkbox(
        "Allow overall QO to exceed +3/-3 in exceptional circumstances (crisis, recent default, sharp recent downgrade, etc.)",
        key="qo_crisis_extension",
        help="Public criteria allow the overall QO range to extend in certain circumstances.",
    )

    raw = {
        "structural": int(st.session_state.get("qo_structural", 0)),
        "macro": int(st.session_state.get("qo_macro", 0)),
        "public_finances": int(st.session_state.get("qo_public_finances", 0)),
        "external": int(st.session_state.get("qo_external", 0)),
    }
    qo_total = clamp_qo(raw, crisis_extension=crisis)

    c1, c2 = st.columns(2)
    c1.metric("Raw QO total", f"{sum(raw.values()):+d}")
    c2.metric("Applied QO total", f"{qo_total:+d}")
    if not crisis and qo_total != sum(raw.values()):
        st.info("The raw QO total has been capped to +3/-3, consistent with Fitch's typical overall QO range.")


def page_results():
    st.title("Results")
    srm_inputs = {k: float(st.session_state.get(k, 0.0)) for pillar in SRM_VARIABLES.values() for k in pillar.keys()}
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
        st.warning(
            "Public Fitch criteria say that when the proposed LT FC IDR is CCC+ or below, ratings are directly based on Fitch's rating definitions rather than explained through SRM + QO. Treat the output below as model guidance only."
        )

    st.markdown("---")
    st.subheader("Long-term LC / FC and short-term ratings")
    st.caption("Fitch generally equates FC and LC IDRs except near distress; this tool uses a user-controlled notch adjustment for LC vs FC.")

    lc_adjust = st.selectbox(
        "Manual LC rating notch difference versus FC",
        options=[-2, -1, 0, 1, 2],
        index=[-2, -1, 0, 1, 2].index(int(st.session_state.get("lc_manual_adjust", 0))),
        key="lc_manual_adjust",
        format_func=lambda x: f"{x:+d} notch(es)",
        help="Positive means LC rated above FC; negative means LC below FC.",
    )
    lt_lc = apply_notches(implied_lt_fc, lc_adjust)

    robust_liquidity = st.checkbox(
        "Use higher short-term FC mapping option (reserve-currency flexibility > 0 or robust external liquidity)",
        key="fc_robust_liquidity",
        help="Per Fitch's short-term mapping guidance, FC ST can use the higher option if reserve-currency flexibility is present or liquidity is robust.",
    )
    st_fc = map_short_term(implied_lt_fc, robust_liquidity)
    st_lc = map_short_term(lt_lc, True)

    r1, r2, r3, r4 = st.columns(4)
    r1.metric("LT FC IDR", implied_lt_fc)
    r2.metric("LT LC IDR", lt_lc)
    r3.metric("ST FC IDR", st_fc)
    r4.metric("ST LC IDR", st_lc)

    st.markdown("---")
    st.subheader("SRM contribution breakdown")
    df = pd.DataFrame(details)
    df = df.sort_values(["Pillar", "Contribution"], ascending=[True, False])
    st.dataframe(df, use_container_width=True, hide_index=True)

    fig = build_radar(srm_score, qo_total, final_model_score)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader("Export snapshot")
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
    }
    st.download_button(
        "Download JSON snapshot",
        data=json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"),
        file_name="fitch_sovereign_methodology_snapshot.json",
        mime="application/json",
    )


def page_instruments():
    st.title("Debt Instruments and Recovery")
    st.caption("Fitch typically aligns senior unsecured sovereign debt with the LT IDR at BB- or above. At B+ and below, recovery notching can apply.")

    lt_base = st.selectbox(
        "Base long-term rating for the instrument analysis",
        options=["B+", "B", "B-", "CCC+", "CCC", "CCC-", "CC", "C/RD"],
        index=0,
        key="instrument_base_rating",
    )
    rr = st.selectbox("Recovery Rating", RR_OPTIONS, index=RR_OPTIONS.index(st.session_state.get("rr_choice", "RR4")), key="rr_choice")

    instrument_rating = debt_instrument_rating(lt_base, rr)
    c1, c2 = st.columns(2)
    c1.metric("Base issuer rating", lt_base)
    c2.metric("Illustrative instrument rating", instrument_rating)

    table_rows = []
    for rr_name in RR_OPTIONS:
        row = {"RR": rr_name}
        for base in ["B+", "B", "B-", "CCC+", "CCC", "CCC-", "CC", "C/RD"]:
            row[base] = debt_instrument_rating(base, rr_name)
        table_rows.append(row)
    st.dataframe(pd.DataFrame(table_rows), use_container_width=True, hide_index=True)

    st.markdown(
        """
### Notes
- This page is meant for **senior unsecured sovereign instruments** where Fitch believes recovery analysis is informative.
- The simplified table follows the public criteria matrix for combinations of **issuer IDRs** and **Recovery Ratings**.
- For higher ratings, sovereign long-term instrument ratings are usually aligned with the issuer rating.
        """
    )


def main():
    init_state()

    st.sidebar.title("Fitch methodology app")
    page = st.sidebar.radio(
        "Navigate",
        [
            "Overview",
            PILLAR_LABELS["structural"],
            PILLAR_LABELS["macro"],
            PILLAR_LABELS["public_finances"],
            PILLAR_LABELS["external"],
            "Qualitative Overlay (QO)",
            "Results",
            "Debt Instruments and Recovery",
        ],
    )

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
    elif page == "Debt Instruments and Recovery":
        page_instruments()


if __name__ == "__main__":
    main()
