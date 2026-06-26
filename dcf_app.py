"""
PepsiCo DCF Valuation Model — Interactive Streamlit App
Built by: Poonam Dhanuka | DePaul MS Finance 2027
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import io
from datetime import datetime

# ── PAGE CONFIG ──────────────────────────────────────────────────
st.set_page_config(
    page_title="PepsiCo DCF Valuation",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── CUSTOM CSS ───────────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background-color: #0f0f0f; color: #e0e0e0; }
#MainMenu, footer { visibility: hidden; }
section[data-testid="stSidebar"] {
    background-color: #111 !important;
    border-right: 1px solid #222 !important;
}
section[data-testid="stSidebar"] * { color: #ccc !important; }

.dcf-header {
    background: linear-gradient(135deg, #1B3A6B 0%, #2d5aa0 100%);
    border-radius: 14px;
    padding: 24px 28px;
    margin-bottom: 20px;
}
.dcf-title { font-size: 26px; font-weight: 800; color: #fff; }
.dcf-sub   { font-size: 13px; color: #aac; margin-top: 4px; }

.result-card {
    background: #141414;
    border: 1px solid #222;
    border-radius: 12px;
    padding: 20px 24px;
    text-align: center;
}
.result-label { font-size: 11px; color: #666; text-transform: uppercase; letter-spacing: 0.8px; }
.result-value { font-size: 32px; font-weight: 800; color: #fff; margin-top: 4px; }
.result-sub   { font-size: 12px; color: #888; margin-top: 2px; }

.metric-card {
    background: #141414;
    border: 1px solid #222;
    border-radius: 10px;
    padding: 14px 16px;
}
.metric-label { font-size: 10px; color: #666; text-transform: uppercase; letter-spacing: 0.5px; }
.metric-value { font-size: 18px; font-weight: 700; color: #fff; margin-top: 3px; }

.section-title {
    font-size: 15px; font-weight: 700; color: #fff;
    border-left: 3px solid #1B3A6B;
    padding-left: 10px; margin: 20px 0 12px 0;
}
.assumption-note {
    background: #1a1a1a; border: 1px solid #2a2a2a;
    border-radius: 8px; padding: 10px 14px;
    font-size: 12px; color: #888; margin-top: 8px;
}
hr { border-color: #222 !important; }

[data-testid="metric-container"] {
    background: #141414; border: 1px solid #222;
    border-radius: 10px; padding: 14px 16px !important;
}
[data-testid="metric-container"] label { color: #888 !important; font-size: 11px !important; }
[data-testid="metric-container"] [data-testid="stMetricValue"] { color: #fff !important; font-weight: 700 !important; }
.stTabs [data-baseweb="tab-list"] { background: #141414; border-radius: 8px; }
.stTabs [data-baseweb="tab"]      { color: #888 !important; }
.stTabs [aria-selected="true"]    { background: #1B3A6B !important; color: #fff !important; }
.stButton > button {
    background: #1B3A6B !important; color: #fff !important;
    border: none !important; border-radius: 8px !important; font-weight: 600 !important;
}
</style>
""", unsafe_allow_html=True)

# ── DARK PLOTLY LAYOUT ───────────────────────────────────────────
DARK = dict(
    paper_bgcolor='#111', plot_bgcolor='#111',
    font=dict(color='#aaa', size=12),
    xaxis=dict(showgrid=False, color='#444', zeroline=False),
    yaxis=dict(gridcolor='#222', color='#aaa', zeroline=False),
    margin=dict(l=10, r=10, t=40, b=10),
    legend=dict(bgcolor='rgba(0,0,0,0)', font=dict(color='#aaa'))
)

# ── SIDEBAR: ASSUMPTIONS ─────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 DCF Assumptions")
    st.markdown('<span style="color:#888;font-size:12px">Adjust any value — model updates instantly</span>', unsafe_allow_html=True)
    st.divider()

    st.markdown("**📈 Revenue Growth Rates**")
    g1 = st.slider("2026 Growth %", 0.0, 15.0, 4.0, 0.5) / 100
    g2 = st.slider("2027 Growth %", 0.0, 15.0, 4.0, 0.5) / 100
    g3 = st.slider("2028 Growth %", 0.0, 15.0, 3.5, 0.5) / 100
    g4 = st.slider("2029 Growth %", 0.0, 15.0, 3.5, 0.5) / 100
    g5 = st.slider("2030 Growth %", 0.0, 15.0, 3.0, 0.5) / 100
    revenue_growth = [g1, g2, g3, g4, g5]

    st.divider()
    st.markdown("**💰 Margin Assumptions**")
    ebit_margin     = st.slider("EBIT Margin %",      5.0, 30.0, 14.5, 0.5) / 100
    tax_rate        = st.slider("Tax Rate %",          10.0, 35.0, 21.0, 1.0) / 100
    da_pct          = st.slider("D&A (% of Revenue)", 1.0, 10.0, 4.5,  0.5) / 100
    capex_pct       = st.slider("CapEx (% of Revenue)", 2.0, 12.0, 5.0, 0.5) / 100
    nwc_pct         = st.slider("NWC Change (% of Revenue)", 0.0, 5.0, 0.5, 0.25) / 100

    st.divider()
    st.markdown("**🏦 WACC Inputs**")
    rf_rate         = st.slider("Risk-Free Rate %",         2.0, 6.0, 4.5, 0.25) / 100
    beta            = st.slider("Beta",                     0.3, 1.5, 0.65, 0.05)
    erp             = st.slider("Equity Risk Premium %",    3.0, 8.0, 5.5, 0.25) / 100
    pretax_kd       = st.slider("Pre-Tax Cost of Debt %",  2.0, 7.0, 3.6, 0.1) / 100
    debt_weight     = st.slider("Debt Weight %",           20.0, 80.0, 55.0, 5.0) / 100
    equity_weight   = 1 - debt_weight

    st.divider()
    st.markdown("**🎯 Terminal Value**")
    ltgr            = st.slider("Long-Term Growth Rate %",  1.0, 4.0, 2.5, 0.25) / 100
    exit_mult       = st.slider("Exit EV/EBITDA Multiple", 8.0, 20.0, 13.0, 0.5)

    st.divider()
    st.markdown("**🏢 Balance Sheet (as of 2025)**")
    gross_debt      = st.number_input("Gross Debt ($mn)",         value=42000, step=500)
    cash            = st.number_input("Cash ($mn)",               value=9000,  step=500)
    non_op_assets   = st.number_input("Non-Operating Assets ($mn)",value=2500, step=100)
    nci             = st.number_input("Non-Controlling Interests ($mn)", value=120, step=10)
    shares_out      = st.number_input("Diluted Shares (mn)",      value=1380,  step=10)

    st.caption("Data: Yahoo Finance (PEP) | DePaul MS Finance 2027")


# ─────────────────────────────────────────────────────────────────
# MODEL ENGINE
# ─────────────────────────────────────────────────────────────────

BASE_REVENUE = 92_000
YEARS        = [2026, 2027, 2028, 2029, 2030]

@st.cache_data
def run_model(revenue_growth, ebit_margin, tax_rate, da_pct, capex_pct,
              nwc_pct, rf_rate, beta, erp, pretax_kd, debt_weight,
              ltgr, exit_mult, gross_debt, cash, non_op_assets, nci, shares_out):

    # ── Projections ───────────────────────────────────────────────
    revenues, ebits, nopats, das, ebitdas, capexs, nwcs, fcffs, taxes = [],[],[],[],[],[],[],[],[]
    for i in range(5):
        rev   = (revenues[-1] if i else BASE_REVENUE) * (1 + revenue_growth[i])
        ebit  = rev * ebit_margin
        tax   = ebit * tax_rate
        nopat = ebit - tax
        da    = rev * da_pct
        capex = rev * capex_pct
        nwc   = rev * nwc_pct
        fcff  = nopat + da - nwc - capex
        revenues.append(round(rev)); ebits.append(round(ebit))
        taxes.append(round(tax));    nopats.append(round(nopat))
        das.append(round(da));       ebitdas.append(round(ebit+da))
        capexs.append(round(capex)); nwcs.append(round(nwc))
        fcffs.append(round(fcff))

    # ── WACC ──────────────────────────────────────────────────────
    ke   = rf_rate + beta * erp
    kd   = pretax_kd * (1 - tax_rate)
    wacc = equity_weight * ke + debt_weight * kd   # Note: equity_weight passed as param

    # ── Discounting (mid-year) ────────────────────────────────────
    pv_f    = [(1/(1+wacc)**(i+0.5)) for i in range(5)]
    pv_fcff = [round(f*p) for f, p in zip(fcffs, pv_f)]
    pv_exp  = sum(pv_fcff)

    # ── Terminal Value: Gordon Growth ─────────────────────────────
    tv_gg     = fcffs[-1] * (1 + ltgr) / (wacc - ltgr)  if wacc > ltgr else 0
    pv_tv_gg  = round(tv_gg * pv_f[-1])
    ev_gg     = pv_exp + pv_tv_gg
    eq_gg     = ev_gg + non_op_assets + cash - gross_debt - nci
    pps_gg    = round(eq_gg / shares_out, 2) if shares_out else 0

    # ── Terminal Value: Exit Multiple ─────────────────────────────
    tv_exit   = ebitdas[-1] * exit_mult
    pv_tv_ex  = round(tv_exit * pv_f[-1])
    ev_exit   = pv_exp + pv_tv_ex
    eq_exit   = ev_exit + non_op_assets + cash - gross_debt - nci
    pps_exit  = round(eq_exit / shares_out, 2) if shares_out else 0

    return dict(
        revenues=revenues, ebits=ebits, taxes=taxes, nopats=nopats,
        das=das, ebitdas=ebitdas, capexs=capexs, nwcs=nwcs, fcffs=fcffs,
        pv_f=[round(p,4) for p in pv_f], pv_fcff=pv_fcff, pv_exp=round(pv_exp),
        ke=round(ke,4), kd=round(kd,4), wacc=round(wacc,4),
        tv_gg=round(tv_gg), pv_tv_gg=pv_tv_gg, ev_gg=round(ev_gg), eq_gg=round(eq_gg), pps_gg=pps_gg,
        tv_exit=round(tv_exit), pv_tv_ex=pv_tv_ex, ev_exit=round(ev_exit), eq_exit=round(eq_exit), pps_exit=pps_exit,
    )

# Run the model
m = run_model(
    tuple(revenue_growth), ebit_margin, tax_rate, da_pct, capex_pct, nwc_pct,
    rf_rate, beta, erp, pretax_kd, debt_weight,
    ltgr, exit_mult, gross_debt, cash, non_op_assets, nci, shares_out
)
equity_weight = 1 - debt_weight


# ─────────────────────────────────────────────────────────────────
# MAIN PAGE
# ─────────────────────────────────────────────────────────────────

# Header
st.markdown(f"""
<div class="dcf-header">
    <div class="dcf-title">📊 PepsiCo DCF Valuation Model</div>
    <div class="dcf-sub">
        Forecast Period: 2026 – 2030 &nbsp;·&nbsp; WACC: {m['wacc']*100:.2f}%
        &nbsp;·&nbsp; Built by Poonam Dhanuka · DePaul MS Finance 2027
    </div>
</div>
""", unsafe_allow_html=True)

# ── RESULT CARDS ────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(f"""<div class="result-card">
        <div class="result-label">Price/Share (Gordon Growth)</div>
        <div class="result-value" style="color:#00c878">${m['pps_gg']:.2f}</div>
        <div class="result-sub">EV: ${m['ev_gg']:,}mn</div>
    </div>""", unsafe_allow_html=True)
with c2:
    st.markdown(f"""<div class="result-card">
        <div class="result-label">Price/Share (Exit Multiple)</div>
        <div class="result-value" style="color:#6001d2">${m['pps_exit']:.2f}</div>
        <div class="result-sub">EV: ${m['ev_exit']:,}mn</div>
    </div>""", unsafe_allow_html=True)
with c3:
    st.markdown(f"""<div class="result-card">
        <div class="result-label">WACC</div>
        <div class="result-value">{m['wacc']*100:.2f}%</div>
        <div class="result-sub">Ke: {m['ke']*100:.2f}% · Kd: {m['kd']*100:.2f}%</div>
    </div>""", unsafe_allow_html=True)
with c4:
    avg_pps = (m['pps_gg'] + m['pps_exit']) / 2
    st.markdown(f"""<div class="result-card">
        <div class="result-label">Blended Fair Value</div>
        <div class="result-value" style="color:#f59e0b">${avg_pps:.2f}</div>
        <div class="result-sub">Average of both methods</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── TABS ────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📈 Projections", "💵 FCF & Discounting", "🏦 WACC", "🎯 Sensitivity", "📥 Export"
])

# ──────────────────────────────────────────────────────────────
# TAB 1: PROJECTIONS
# ──────────────────────────────────────────────────────────────
with tab1:
    col1, col2 = st.columns(2)

    with col1:
        st.markdown('<div class="section-title">Revenue & EBIT (2026–2030)</div>', unsafe_allow_html=True)
        fig = go.Figure()
        fig.add_trace(go.Bar(x=YEARS, y=m['revenues'], name='Revenue',
                             marker_color='#1B3A6B', opacity=0.85))
        fig.add_trace(go.Bar(x=YEARS, y=m['ebits'], name='EBIT',
                             marker_color='#6001d2', opacity=0.85))
        fig.update_layout(barmode='group', height=320, **DARK,
                          title="Revenue vs EBIT ($mn)")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.markdown('<div class="section-title">EBITDA & FCFF</div>', unsafe_allow_html=True)
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(x=YEARS, y=m['ebitdas'], name='EBITDA',
                              marker_color='#2d5aa0', opacity=0.85))
        fig2.add_trace(go.Scatter(x=YEARS, y=m['fcffs'], name='FCFF',
                                  line=dict(color='#00c878', width=3),
                                  mode='lines+markers',
                                  marker=dict(size=8, color='#00c878')))
        fig2.update_layout(height=320, **DARK, title="EBITDA vs FCFF ($mn)")
        st.plotly_chart(fig2, use_container_width=True)

    # Table
    st.markdown('<div class="section-title">Projected Income Statement ($mn)</div>', unsafe_allow_html=True)
    df_proj = pd.DataFrame({
        "Year":         YEARS,
        "Revenue":      [f"${v:,}" for v in m['revenues']],
        "EBIT":         [f"${v:,}" for v in m['ebits']],
        "EBIT Margin":  [f"{v/r*100:.1f}%" for v,r in zip(m['ebits'], m['revenues'])],
        "Tax":          [f"${v:,}" for v in m['taxes']],
        "NOPAT":        [f"${v:,}" for v in m['nopats']],
        "D&A":          [f"${v:,}" for v in m['das']],
        "EBITDA":       [f"${v:,}" for v in m['ebitdas']],
        "CapEx":        [f"(${v:,})" for v in m['capexs']],
        "FCFF":         [f"${v:,}" for v in m['fcffs']],
    })
    st.dataframe(df_proj.set_index("Year"), use_container_width=True)

# ──────────────────────────────────────────────────────────────
# TAB 2: FCF & DISCOUNTING
# ──────────────────────────────────────────────────────────────
with tab2:
    col1, col2 = st.columns(2)

    with col1:
        st.markdown('<div class="section-title">PV of FCFFs vs Terminal Value</div>', unsafe_allow_html=True)
        labels = [str(y) for y in YEARS] + ["Terminal Value\n(GG)"]
        values_gg = m['pv_fcff'] + [m['pv_tv_gg']]
        colors = ['#1B3A6B']*5 + ['#00c878']
        fig3 = go.Figure(go.Bar(x=labels, y=values_gg, marker_color=colors, opacity=0.9))
        fig3.update_layout(height=320, **DARK,
                           title="PV of Cash Flows — Gordon Growth ($mn)")
        st.plotly_chart(fig3, use_container_width=True)

    with col2:
        st.markdown('<div class="section-title">Enterprise Value Waterfall</div>', unsafe_allow_html=True)
        fig4 = go.Figure(go.Waterfall(
            name="EV Bridge", orientation="v",
            measure=["relative","relative","total","relative","relative","total"],
            x=["PV Explicit FCFs","PV Terminal Value (GG)","Enterprise Value",
               "+ Cash & Non-Op","- Gross Debt","Equity Value"],
            y=[m['pv_exp'], m['pv_tv_gg'], 0,
               cash + non_op_assets, -gross_debt, 0],
            connector=dict(line=dict(color="#444")),
            increasing=dict(marker=dict(color="#00c878")),
            decreasing=dict(marker=dict(color="#ff4b4b")),
            totals=dict(marker=dict(color="#1B3A6B")),
        ))
        fig4.update_layout(height=320, **DARK,
                           title="Equity Value Bridge — Gordon Growth ($mn)")
        st.plotly_chart(fig4, use_container_width=True)

    # Discounting table
    st.markdown('<div class="section-title">Discounting Schedule</div>', unsafe_allow_html=True)
    df_disc = pd.DataFrame({
        "Year":         YEARS,
        "FCFF ($mn)":   [f"${v:,}" for v in m['fcffs']],
        "PV Factor":    [f"{v:.4f}" for v in m['pv_f']],
        "PV of FCFF":   [f"${v:,}" for v in m['pv_fcff']],
    })
    st.dataframe(df_disc.set_index("Year"), use_container_width=True)

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown('<div class="section-title">Gordon Growth Method</div>', unsafe_allow_html=True)
        st.metric("PV of Explicit FCFs",      f"${m['pv_exp']:,}mn")
        st.metric("Terminal Value",            f"${m['tv_gg']:,}mn")
        st.metric("PV of Terminal Value",      f"${m['pv_tv_gg']:,}mn")
        st.metric("Enterprise Value",          f"${m['ev_gg']:,}mn")
        st.metric("Equity Value",              f"${m['eq_gg']:,}mn")
        st.metric("★ Price per Share",         f"${m['pps_gg']:.2f}")
    with col_b:
        st.markdown('<div class="section-title">Exit Multiple Method</div>', unsafe_allow_html=True)
        st.metric("PV of Explicit FCFs",       f"${m['pv_exp']:,}mn")
        st.metric(f"Terminal Value ({exit_mult}x EBITDA)", f"${m['tv_exit']:,}mn")
        st.metric("PV of Terminal Value",       f"${m['pv_tv_ex']:,}mn")
        st.metric("Enterprise Value",           f"${m['ev_exit']:,}mn")
        st.metric("Equity Value",               f"${m['eq_exit']:,}mn")
        st.metric("★ Price per Share",          f"${m['pps_exit']:.2f}")

# ──────────────────────────────────────────────────────────────
# TAB 3: WACC
# ──────────────────────────────────────────────────────────────
with tab3:
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="section-title">Cost of Equity (CAPM)</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card"><div class="metric-label">Risk-Free Rate (Rf)</div><div class="metric-value">{rf_rate*100:.2f}%</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card"><div class="metric-label">Beta (β)</div><div class="metric-value">{beta:.2f}</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card"><div class="metric-label">Equity Risk Premium (ERP)</div><div class="metric-value">{erp*100:.2f}%</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card" style="border-color:#00c878"><div class="metric-label">Cost of Equity = Rf + β×ERP</div><div class="metric-value" style="color:#00c878">{m["ke"]*100:.2f}%</div></div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="section-title">Cost of Debt</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card"><div class="metric-label">Pre-Tax Cost of Debt</div><div class="metric-value">{pretax_kd*100:.2f}%</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card"><div class="metric-label">Tax Rate</div><div class="metric-value">{tax_rate*100:.0f}%</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card" style="border-color:#6001d2"><div class="metric-label">After-Tax Cost of Debt</div><div class="metric-value" style="color:#6001d2">{m["kd"]*100:.2f}%</div></div>', unsafe_allow_html=True)

    with col3:
        st.markdown('<div class="section-title">Capital Structure & WACC</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card"><div class="metric-label">Debt Weight</div><div class="metric-value">{debt_weight*100:.0f}%</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card"><div class="metric-label">Equity Weight</div><div class="metric-value">{equity_weight*100:.0f}%</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card" style="border-color:#f59e0b"><div class="metric-label">WACC = Wd×Kd + We×Ke</div><div class="metric-value" style="color:#f59e0b">{m["wacc"]*100:.2f}%</div></div>', unsafe_allow_html=True)

    # WACC pie
    fig_w = go.Figure(go.Pie(
        labels=['Debt', 'Equity'],
        values=[debt_weight*100, equity_weight*100],
        hole=0.5,
        marker_colors=['#6001d2', '#1B3A6B'],
    ))
    fig_w.update_layout(height=280, **DARK, title="Capital Structure")
    st.plotly_chart(fig_w, use_container_width=True)

# ──────────────────────────────────────────────────────────────
# TAB 4: SENSITIVITY
# ──────────────────────────────────────────────────────────────
with tab4:
    st.markdown('<div class="section-title">Sensitivity: WACC vs Long-Term Growth Rate → Implied Share Price ($)</div>', unsafe_allow_html=True)
    st.markdown("🟡 = Base case &nbsp;&nbsp; 🟢 = >10% above base &nbsp;&nbsp; 🔴 = >10% below base", unsafe_allow_html=True)

    wacc_range = [round(m['wacc'] + d, 4) for d in [-0.015,-0.010,-0.005,0,0.005,0.010,0.015]]
    ltgr_range = [0.015, 0.020, 0.025, 0.030, 0.035]
    base_pps   = m['pps_gg']

    rows = {}
    for w in wacc_range:
        row = {}
        pv_f2  = [(1/(1+w)**(i+0.5)) for i in range(5)]
        pv_ex2 = sum(f*p for f,p in zip(m['fcffs'], pv_f2))
        for g in ltgr_range:
            col_name = f"{g*100:.1f}%"
            if w <= g:
                row[col_name] = None
            else:
                tv2  = m['fcffs'][-1]*(1+g)/(w-g)
                pv2  = tv2 * pv_f2[-1]
                ev2  = pv_ex2 + pv2
                eq2  = ev2 + non_op_assets + cash - gross_debt - nci
                row[col_name] = round(eq2/shares_out, 2)
        rows[f"{w*100:.1f}%"] = row

    df_sens = pd.DataFrame(rows).T
    df_sens.index.name = "WACC \\ LTGR"

    # Style: color cells
    def color_cell(val):
        if val is None: return 'background-color:#1a1a1a; color:#444'
        if val > base_pps * 1.10: return 'background-color:#1A5C34; color:#fff; font-weight:bold'
        if val < base_pps * 0.90: return 'background-color:#5C1A1A; color:#fff; font-weight:bold'
        return 'background-color:#1a2a3a; color:#eee'

    base_wacc_label = f"{m['wacc']*100:.1f}%"
    base_ltgr_label = f"{ltgr*100:.1f}%"

    styled = df_sens.style.map(color_cell).format(
        lambda x: f"${x:.2f}" if x is not None else "N/A"
    )
    st.dataframe(styled, use_container_width=True, height=320)

    # Tornado chart
    st.markdown('<div class="section-title">Impact on Share Price (Gordon Growth)</div>', unsafe_allow_html=True)
    vars_impact = {
        "Revenue Growth +1%":    run_model(tuple([g+0.01 for g in revenue_growth]), ebit_margin, tax_rate, da_pct, capex_pct, nwc_pct, rf_rate, beta, erp, pretax_kd, debt_weight, ltgr, exit_mult, gross_debt, cash, non_op_assets, nci, shares_out)['pps_gg'] - base_pps,
        "EBIT Margin +1%":       run_model(tuple(revenue_growth), ebit_margin+0.01, tax_rate, da_pct, capex_pct, nwc_pct, rf_rate, beta, erp, pretax_kd, debt_weight, ltgr, exit_mult, gross_debt, cash, non_op_assets, nci, shares_out)['pps_gg'] - base_pps,
        "WACC -0.5%":            run_model(tuple(revenue_growth), ebit_margin, tax_rate, da_pct, capex_pct, nwc_pct, rf_rate, beta, erp, pretax_kd, debt_weight, ltgr-0.005, exit_mult, gross_debt, cash, non_op_assets, nci, shares_out)['pps_gg'] - base_pps,
        "LTGR +0.5%":            run_model(tuple(revenue_growth), ebit_margin, tax_rate, da_pct, capex_pct, nwc_pct, rf_rate, beta, erp, pretax_kd, debt_weight, ltgr+0.005, exit_mult, gross_debt, cash, non_op_assets, nci, shares_out)['pps_gg'] - base_pps,
        "CapEx -1% Rev":         run_model(tuple(revenue_growth), ebit_margin, tax_rate, da_pct, capex_pct-0.01, nwc_pct, rf_rate, beta, erp, pretax_kd, debt_weight, ltgr, exit_mult, gross_debt, cash, non_op_assets, nci, shares_out)['pps_gg'] - base_pps,
    }
    fig_t = go.Figure(go.Bar(
        x=list(vars_impact.values()),
        y=list(vars_impact.keys()),
        orientation='h',
        marker_color=['#00c878' if v>=0 else '#ff4b4b' for v in vars_impact.values()],
        opacity=0.9,
    ))
    fig_t.update_layout(height=280, **DARK, title="Sensitivity Impact on Share Price ($)")
    fig_t.add_vline(x=0, line_dash="dash", line_color="#888")
    st.plotly_chart(fig_t, use_container_width=True)

# ──────────────────────────────────────────────────────────────
# TAB 5: EXPORT
# ──────────────────────────────────────────────────────────────
with tab5:
    st.markdown('<div class="section-title">Download Excel Report</div>', unsafe_allow_html=True)
    st.markdown("Click below to download a formatted Excel workbook with all assumptions, projections, and results.")

    def build_excel():
        wb = Workbook()
        ws = wb.active
        ws.title = "DCF Summary"
        ws.sheet_view.showGridLines = False

        NAVY_F  = PatternFill("solid", fgColor="1B3A6B")
        LBLUE_F = PatternFill("solid", fgColor="D6E4FF")
        GREEN_F = PatternFill("solid", fgColor="D5F0E3")
        WHITE_F = PatternFill("solid", fgColor="FFFFFF")

        def hdr(row, col, text, bg=NAVY_F, fg="FFFFFF", bold=True, merge=None):
            c = ws.cell(row=row, column=col, value=text)
            c.fill = bg; c.font = Font(color=fg, bold=bold, name="Calibri", size=10)
            c.alignment = Alignment(horizontal="center", vertical="center")
            if merge: ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=merge)

        def val(row, col, v, fmt="#,##0", bold=False, bg=None):
            c = ws.cell(row=row, column=col, value=v)
            c.number_format = fmt
            c.font = Font(bold=bold, name="Calibri")
            c.alignment = Alignment(horizontal="right", vertical="center")
            if bg: c.fill = bg

        ws.column_dimensions["A"].width = 2
        ws.column_dimensions["B"].width = 40
        for col in ["C","D","E","F","G"]: ws.column_dimensions[col].width = 14

        hdr(1, 2, "PepsiCo DCF Valuation — 2026 to 2030", merge=7)
        hdr(2, 2, f"Built by Poonam Dhanuka | DePaul MS Finance 2027 | {datetime.today().strftime('%B %d, %Y')}", bg=LBLUE_F, fg="1B3A6B", bold=False, merge=7)

        hdr(4, 2, "Particulars")
        for i, y in enumerate(YEARS): hdr(4, 3+i, str(y))

        rows = [
            ("Revenue",       m['revenues']),
            ("EBIT",          m['ebits']),
            ("(-) Tax",       m['taxes']),
            ("NOPAT",         m['nopats']),
            ("D&A",           m['das']),
            ("(-) CapEx",     m['capexs']),
            ("FCFF",          m['fcffs']),
            ("PV Factor",     m['pv_f']),
            ("PV of FCFF",    m['pv_fcff']),
        ]
        fmts = ["#,##0"]*7 + ["0.0000","#,##0"]
        for ri, (label, values) in enumerate(rows):
            r = 5 + ri
            bold = label in ("NOPAT","FCFF","PV of FCFF")
            bg   = GREEN_F if bold else None
            ws.cell(row=r, column=2, value=label).font = Font(bold=bold, name="Calibri")
            for ci, v in enumerate(values): val(r, 3+ci, v, fmts[ri], bold=bold, bg=bg)

        r = 15
        hdr(r, 2, "VALUATION SUMMARY", merge=5)
        summary = [
            ("WACC",                          f"{m['wacc']*100:.2f}%"),
            ("PV of Explicit FCFs ($mn)",     m['pv_exp']),
            ("Enterprise Value — GG ($mn)",   m['ev_gg']),
            ("Equity Value — GG ($mn)",       m['eq_gg']),
            ("Price per Share — GG ($)",      m['pps_gg']),
            ("Enterprise Value — Exit ($mn)", m['ev_exit']),
            ("Equity Value — Exit ($mn)",     m['eq_exit']),
            ("Price per Share — Exit ($)",    m['pps_exit']),
        ]
        for i, (label, v) in enumerate(summary):
            r2 = 16 + i
            ws.cell(row=r2, column=2, value=label).font = Font(name="Calibri")
            c = ws.cell(row=r2, column=3, value=v)
            c.font = Font(bold="Price" in label, name="Calibri")
            c.alignment = Alignment(horizontal="right")
            if "Price" in label: c.fill = GREEN_F

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    excel_data = build_excel()
    st.download_button(
        label="📥 Download Excel Valuation Report",
        data=excel_data,
        file_name=f"PepsiCo_DCF_{datetime.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.markdown('<div class="section-title">Current Assumptions Summary</div>', unsafe_allow_html=True)
    assumptions_df = pd.DataFrame({
        "Parameter": ["Base Revenue ($mn)","Revenue Growth (2026-2030)","EBIT Margin",
                      "Tax Rate","D&A % Revenue","CapEx % Revenue","NWC Change % Revenue",
                      "Risk-Free Rate","Beta","ERP","WACC",
                      "Long-Term Growth Rate","Exit EV/EBITDA Multiple",
                      "Gross Debt ($mn)","Cash ($mn)","Shares Outstanding (mn)"],
        "Value": [f"${BASE_REVENUE:,}",
                  " / ".join([f"{g*100:.1f}%" for g in revenue_growth]),
                  f"{ebit_margin*100:.1f}%", f"{tax_rate*100:.0f}%",
                  f"{da_pct*100:.1f}%", f"{capex_pct*100:.1f}%",
                  f"{nwc_pct*100:.2f}%", f"{rf_rate*100:.2f}%",
                  f"{beta:.2f}", f"{erp*100:.1f}%",
                  f"{m['wacc']*100:.2f}%", f"{ltgr*100:.1f}%",
                  f"{exit_mult:.1f}x", f"${gross_debt:,}",
                  f"${cash:,}", f"{shares_out:,}"],
    })
    st.dataframe(assumptions_df.set_index("Parameter"), use_container_width=True)

# ── Footer ───────────────────────────────────────────────────────
st.markdown('<hr>', unsafe_allow_html=True)
st.markdown("""
<div style="text-align:center;color:#444;font-size:11px;padding:8px">
    Built by <b style="color:#1B3A6B">Poonam Dhanuka</b> · DePaul MS Finance 2027 ·
    PepsiCo DCF Model 2026–2030 · For educational purposes only · Not investment advice
</div>
""", unsafe_allow_html=True)
