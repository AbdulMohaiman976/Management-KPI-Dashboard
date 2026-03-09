import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(page_title="Management KPIs Dashboard", page_icon="📊", layout="wide")

st.markdown("""
<style>
    body { background-color: #f5f5f5; }
    .main { background-color: #f5f5f5; }
    .header-title {
        background: linear-gradient(90deg, #1a3a52 0%, #2c5aa0 100%);
        color: white; padding: 20px; border-radius: 8px; text-align: center;
        margin-bottom: 20px; font-weight: bold; font-size: 28px;
    }
    .filter-box {
        background-color: white; padding: 15px; border-radius: 8px;
        margin-bottom: 20px; border: 1px solid #e0e0e0;
    }
    .filter-label {
        font-size: 11px; font-weight: bold; color: #1a3a52;
        text-transform: uppercase; margin-bottom: 4px;
    }
    .kpi-card {
        background-color: white; padding: 20px; border-radius: 6px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1); text-align: center;
        margin-bottom: 12px; border-top: 4px solid #2c5aa0;
        display: flex; flex-direction: column; justify-content: center; align-items: center;
    }
    .kpi-value { font-size: 28px; font-weight: bold; color: #1a3a52; line-height: 1.2; margin: 10px 0; }
    .kpi-label { font-size: 10px; color: #999; text-transform: uppercase; margin: 0; }
    .chart-box {
        background-color: white; padding: 15px; border-radius: 6px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 15px;
    }
    .chart-title { font-size: 12px; font-weight: bold; color: #1a3a52; margin-bottom: 10px; text-transform: uppercase; }
    .section-header {
        font-size: 16px; font-weight: bold; color: #1a3a52;
        padding: 8px 0; border-bottom: 2px solid #2c5aa0; margin-bottom: 12px;
    }
</style>
""", unsafe_allow_html=True)

# ============ SESSION STATE ============
if 'sales_data' not in st.session_state:
    st.session_state.sales_data = None
if 'financial_data' not in st.session_state:
    st.session_state.financial_data = None
if 'evaluation_data' not in st.session_state:
    st.session_state.evaluation_data = None

# ============ SIDEBAR ============
with st.sidebar:
    st.markdown("### 📁 Upload Excel Files")
    st.markdown("---")
    sales_file = st.file_uploader("1️⃣ Sales KPI", type=['xlsx', 'xls'], key='sales')
    financial_file = st.file_uploader("2️⃣ Financial KPI", type=['xlsx', 'xls'], key='financial')
    evaluation_file = st.file_uploader("3️⃣ Evaluation Form", type=['xlsx', 'xls'], key='evaluation')

    if sales_file:
        st.session_state.sales_data = pd.read_excel(sales_file, sheet_name="Sheet2")
    if financial_file:
        st.session_state.financial_data = pd.read_excel(financial_file, sheet_name="Sheet1")
    if evaluation_file:
        st.session_state.evaluation_data = pd.read_excel(evaluation_file, sheet_name=0)

if (st.session_state.sales_data is None or
    st.session_state.financial_data is None or
    st.session_state.evaluation_data is None):
    st.markdown('<div class="header-title">Management KPIs Dashboard</div>', unsafe_allow_html=True)
    st.info("📁 Upload all 3 Excel files to start")
    st.stop()

# ============ LOAD DATA ============
sales_df = st.session_state.sales_data.copy()
financial_df = st.session_state.financial_data.copy()
evaluation_df = st.session_state.evaluation_data.copy()

sales_df.columns = [' '.join(str(c).split()) for c in sales_df.columns]
financial_df.columns = [' '.join(str(c).split()) for c in financial_df.columns]
evaluation_df.columns = [' '.join(str(c).split()) for c in evaluation_df.columns]

# ============ COLUMN NAMES ============
col_s_dials = "Dials"
col_s_calls = "Calls"
col_s_dm = "DM Conducted"
col_s_dw = "DW Conducted"
col_s_prop_sent = "Proposals Sent"
col_s_prop_sold = "Proposals Sold"
col_s_d2c = "Dials to calls %"
col_s_c2dm = "Calls to DM %"
col_s_dm2dw = "DMs to DW %"
col_s_dw2c = "DWs to Contract %"

col_f_cashflow = "Cashflow_Coverage_Months"
col_f_sales = "Total_Sales_Value"
col_f_cogs = "Cost_of_Goods_Sold"
col_f_total_cost = "Total_Cost"
col_f_net_profit = "Net_Profit_Loss"
col_f_overdue_count = "Overdue_Invoices_Count"
col_f_overdue_val = "Overdue_Invoices_Value"

col_e_trainer = "Trainer Name اسم المدرب"
col_e_content = "How would you rate the content? تقييم المحتوى التدريبي"
col_e_exercise = "How would you rate the exercises? تقييم التمارين التدريبية"
col_e_facilitator = "How would you rate the facilitator? تقييم المدرب"
col_e_expectation = "how did this session compare with your expectations? كيف كانت هذه الجلسة مقارنة بتوقعاتك؟"
col_e_date = "Training Date تاريخ الدورة التدريبية"

# ============ HELPERS ============
def safe_sum(df, col):
    try:
        if col in df.columns:
            return float(pd.to_numeric(df[col], errors='coerce').sum())
        return 0
    except: return 0

def safe_avg(df, col):
    try:
        if col in df.columns:
            return float(pd.to_numeric(df[col], errors='coerce').mean())
        return 0
    except: return 0

MONTH_ORDER = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
               'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12}

def sort_months(lst):
    return sorted(lst, key=lambda m: MONTH_ORDER.get(str(m).strip().lower()[:3], 99))

def val_to_month_str(val):
    if pd.isna(val): return None
    if isinstance(val, str):
        v = val.strip()
        if v[:3].lower() in MONTH_ORDER: return v[:3].capitalize()
        return v
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime('%b')
    try:
        n = int(val)
        if 1 <= n <= 12: return datetime(2024, n, 1).strftime('%b')
    except: pass
    return str(val).strip()

def find_col(df, target):
    t = target.strip().lower()
    for c in df.columns:
        if str(c).strip().lower() == t: return c
    for c in df.columns:
        if t in str(c).strip().lower(): return c
    if t == "month":
        for c in df.columns:
            sample = df[c].dropna().head(10)
            hits = sum(1 for v in sample if val_to_month_str(v) and
                       str(val_to_month_str(v))[:3].lower() in MONTH_ORDER)
            if hits >= max(1, len(sample) * 0.5): return c
    return None

# ============ DETECT COLUMNS ============
actual_s_month = find_col(sales_df, "Month")
actual_f_month = find_col(financial_df, "Month")

# ============ BUILD MONTH LISTS ============
def build_month_list(df, col):
    if not col: return ["All"]
    raw = []
    for v in df[col].dropna().unique():
        m = val_to_month_str(v)
        if m and m.strip() not in ("", "nan"):
            raw.append(m.strip())
    raw = list(dict.fromkeys(raw))
    return ["All"] + sort_months(raw) if raw else ["All"]

sales_months_list   = build_month_list(sales_df, actual_s_month)
finance_months_list = build_month_list(financial_df, actual_f_month)

trainers_list = ["All"]
if col_e_trainer in evaluation_df.columns:
    trainers_list = ["All"] + sorted([str(x).strip() for x in evaluation_df[col_e_trainer].dropna().unique()])

# ============ HEADER ============
st.markdown('<div class="header-title">📊 Management KPIs Dashboard</div>', unsafe_allow_html=True)

# ============ FILTER BAR — ALL FILTERS VISIBLE WITH LABELS ============
st.markdown('<div class="filter-box">', unsafe_allow_html=True)
fc1, fc2, fc3 = st.columns(3)

with fc1:
    st.markdown('<div class="filter-label">📅 Sales Month</div>', unsafe_allow_html=True)
    sel_month_sales = st.selectbox(
        "Sales Month", sales_months_list,
        label_visibility="collapsed", key="sales_mth"
    )

with fc2:
    st.markdown('<div class="filter-label">🎓 Trainer</div>', unsafe_allow_html=True)
    sel_trainer = st.selectbox(
        "Trainer", trainers_list,
        label_visibility="collapsed", key="trainer_sel"
    )

with fc3:
    st.markdown('<div class="filter-label">💰 Finance Month</div>', unsafe_allow_html=True)
    sel_month_finance = st.selectbox(
        "Finance Month", finance_months_list,
        label_visibility="collapsed", key="finance_mth"
    )

st.markdown('</div>', unsafe_allow_html=True)

# ============ FILTER DATA ============
# Sales filter
filt_sales = sales_df.copy()
if sel_month_sales != "All" and actual_s_month:
    filt_sales['_mn'] = filt_sales[actual_s_month].apply(
        lambda v: val_to_month_str(v).strip().lower() if val_to_month_str(v) else ""
    )
    filt_sales = filt_sales[filt_sales['_mn'] == sel_month_sales.strip().lower()]

# Finance filter
filt_fin = financial_df.copy()
if sel_month_finance != "All" and actual_f_month:
    filt_fin['_mn'] = filt_fin[actual_f_month].apply(
        lambda v: val_to_month_str(v).strip().lower() if val_to_month_str(v) else ""
    )
    filt_fin = filt_fin[filt_fin['_mn'] == sel_month_finance.strip().lower()]

# Evaluation filter
filt_eval = evaluation_df.copy()
if sel_trainer != "All" and col_e_trainer in evaluation_df.columns:
    filt_eval = filt_eval[filt_eval[col_e_trainer].astype(str).str.strip() == str(sel_trainer).strip()]

# ============ 3-COLUMN LAYOUT ============
col1, col2, col3 = st.columns(3, gap="small")

# ===== LEFT: SALES KPI =====
with col1:
    st.markdown('<div class="section-header">📈 Sales KPIs</div>', unsafe_allow_html=True)
    st.caption(f"Filter: **{sel_month_sales}**")

    dials     = safe_sum(filt_sales, col_s_dials)
    calls     = safe_sum(filt_sales, col_s_calls)
    dm        = safe_sum(filt_sales, col_s_dm)
    dw        = safe_sum(filt_sales, col_s_dw)
    prop_sent = safe_sum(filt_sales, col_s_prop_sent)
    prop_sold = safe_sum(filt_sales, col_s_prop_sold)

    kc1, kc2, kc3 = st.columns(3)
    with kc1: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{int(dials)}</div><div class="kpi-label">Total Dials</div></div>', unsafe_allow_html=True)
    with kc2: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{int(calls)}</div><div class="kpi-label">Total Calls</div></div>', unsafe_allow_html=True)
    with kc3: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{int(prop_sent)}</div><div class="kpi-label">Proposals Sent</div></div>', unsafe_allow_html=True)

    kc4, kc5, kc6 = st.columns(3)
    with kc4: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{int(dm)}</div><div class="kpi-label">DM Conducted</div></div>', unsafe_allow_html=True)
    with kc5: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{int(dw)}</div><div class="kpi-label">DW Conducted</div></div>', unsafe_allow_html=True)
    with kc6: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{int(prop_sold)}</div><div class="kpi-label">Proposals Sold</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="chart-box"><div class="chart-title">Summary Funnel</div>', unsafe_allow_html=True)
    fig1 = go.Figure(go.Funnel(
        y=['Dials','Calls','DM','DW','Prop Sent','Prop Sold'],
        x=[dials, calls, dm, dw, prop_sent, prop_sold],
        textinfo="value+label", textposition="inside", insidetextanchor="middle",
        marker=dict(color=['#1a3a52','#2c5aa0','#366599','#266092','#4b7eb1','#5e91bd'])
    ))
    fig1.update_layout(height=300, margin=dict(l=20,r=20,t=0,b=20), font=dict(color="white", size=14))
    st.plotly_chart(fig1, use_container_width=True, config={'displayModeBar': False})
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="chart-box"><div class="chart-title">Monthly Sales Conversion Trends</div>', unsafe_allow_html=True)
    if actual_s_month and actual_s_month in filt_sales.columns:
        conv_cols = {}
        for cc in [col_s_d2c, col_s_c2dm, col_s_dm2dw, col_s_dw2c]:
            if cc in filt_sales.columns:
                conv_cols[cc] = lambda x: pd.to_numeric(x, errors='coerce').mean() * 100
        if conv_cols:
            monthly_data = filt_sales.groupby(actual_s_month).agg(conv_cols).fillna(0)
            fig2 = go.Figure()
            for c, name, color in zip(
                [col_s_d2c, col_s_c2dm, col_s_dm2dw, col_s_dw2c],
                ['Dials→Calls %','Calls→DM %','DM→DW %','DW→Contract %'],
                ['#1a3a52','#2c5aa0','#5a8cc7','#8bb3d6']
            ):
                if c in monthly_data.columns:
                    fig2.add_trace(go.Bar(
                        name=name, x=monthly_data.index, y=monthly_data[c],
                        marker_color=color,
                        text=[f"{v:.1f}%" for v in monthly_data[c]],
                        textposition='auto', textfont=dict(color='white', size=12)
                    ))
            fig2.update_layout(height=280, barmode='group', margin=dict(l=20,r=20,t=0,b=30), yaxis_range=[0,100])
            st.plotly_chart(fig2, use_container_width=True, config={'displayModeBar': False})
    st.markdown('</div>', unsafe_allow_html=True)

# ===== MIDDLE: EVALUATION =====
with col2:
    st.markdown('<div class="section-header">🎓 Training Evaluation</div>', unsafe_allow_html=True)
    st.caption(f"Trainer: **{sel_trainer}**")

    st.markdown('<div class="chart-box"><div class="chart-title">Low Expectation over Sessions</div>', unsafe_allow_html=True)
    if col_e_expectation in filt_eval.columns:
        exp_vals = pd.to_numeric(filt_eval[col_e_expectation], errors='coerce').dropna()
        low  = len(exp_vals[exp_vals <= 3])
        good = len(exp_vals[exp_vals > 3])
        fig3 = go.Figure(data=[go.Pie(
            labels=['Low (≤3)','Good (>3)'], values=[low, good],
            marker=dict(colors=["#2c78c5","#4498b9"]),
            textinfo='label+percent', textfont=dict(color='white', size=12)
        )])
        fig3.update_layout(height=250, margin=dict(l=0,r=0,t=0,b=0))
        st.plotly_chart(fig3, use_container_width=True, config={'displayModeBar': False})
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="chart-box"><div class="chart-title">Trainer Performance Comparison</div>', unsafe_allow_html=True)
    perf_cols = {
        col_e_content: 'Content', col_e_exercise: 'Exercise',
        col_e_facilitator: 'Facilitator', col_e_expectation: 'Expectation'
    }
    aggs = {k: (lambda x: pd.to_numeric(x, errors='coerce').mean())
            for k in perf_cols if k in evaluation_df.columns}
    if aggs and col_e_trainer in evaluation_df.columns:
        trainer_perf = evaluation_df.groupby(col_e_trainer).agg(aggs).fillna(0)
        fig4 = go.Figure()
        for k, name, color in zip(aggs.keys(), perf_cols.values(), ['#1a3a52','#2c5aa0','#5a8cc7','#8bb3d6']):
            fig4.add_trace(go.Bar(x=trainer_perf.index, y=trainer_perf[k], name=name, marker_color=color))
        fig4.update_layout(height=320, barmode='group', margin=dict(l=20,r=20,t=0,b=100), xaxis_tickangle=-45)
        st.plotly_chart(fig4, use_container_width=True, config={'displayModeBar': False})
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="chart-box"><div class="chart-title">Month-by-Month Trend Analysis</div>', unsafe_allow_html=True)
    try:
        if col_e_date in evaluation_df.columns:
            edf = evaluation_df.copy()
            edf['date_dt'] = pd.to_datetime(edf[col_e_date], errors='coerce')
            edf['month_name'] = edf['date_dt'].dt.strftime('%b')
            edf['month_num']  = edf['date_dt'].dt.month
            trend_aggs = {k: (lambda x: pd.to_numeric(x, errors='coerce').mean())
                          for k in perf_cols if k in evaluation_df.columns}
            if trend_aggs:
                mt = edf.groupby(['month_num','month_name']).agg(trend_aggs).reset_index().sort_values('month_num')
                fig5 = go.Figure()
                for k, name, color in zip(trend_aggs.keys(), perf_cols.values(), ['#1a3a52','#2c5aa0','#5a8cc7','#8bb3d6']):
                    fig5.add_trace(go.Scatter(x=mt['month_name'], y=mt[k], mode='lines+markers', name=name, line=dict(color=color, width=3)))
                fig5.update_layout(height=280, margin=dict(l=20,r=20,t=0,b=30), hovermode='x unified')
                st.plotly_chart(fig5, use_container_width=True, config={'displayModeBar': False})
    except:
        st.info("Trend unavailable")
    st.markdown('</div>', unsafe_allow_html=True)

# ===== RIGHT: FINANCIAL KPI =====
with col3:
    st.markdown('<div class="section-header">💰 Financial KPIs</div>', unsafe_allow_html=True)
    st.caption(f"Filter: **{sel_month_finance}**")

    cf  = safe_avg(filt_fin, col_f_cashflow)
    sl  = safe_sum(filt_fin, col_f_sales)
    cg  = safe_sum(filt_fin, col_f_cogs)
    np_ = safe_sum(filt_fin, col_f_net_profit)
    oc  = safe_sum(filt_fin, col_f_overdue_count)
    ov  = safe_sum(filt_fin, col_f_overdue_val)

    def fmt_k(v): return f"{int(v/1000)}K" if v >= 1000 else str(int(v))

    kf1, kf2, kf3 = st.columns(3)
    with kf1: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{cf:.1f}</div><div class="kpi-label">Cashflow</div></div>', unsafe_allow_html=True)
    with kf2: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{fmt_k(sl)}</div><div class="kpi-label">Total Sales</div></div>', unsafe_allow_html=True)
    with kf3: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{fmt_k(cg)}</div><div class="kpi-label">Total COGS</div></div>', unsafe_allow_html=True)

    kf4, kf5, kf6 = st.columns(3)
    with kf4: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{fmt_k(np_)}</div><div class="kpi-label">Net Profit</div></div>', unsafe_allow_html=True)
    with kf5: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{int(oc)}</div><div class="kpi-label">Overdue Count</div></div>', unsafe_allow_html=True)
    with kf6: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{fmt_k(ov)}</div><div class="kpi-label">Overdue Val</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="chart-box"><div class="chart-title">Net Profit/Loss Over Month</div>', unsafe_allow_html=True)
    if actual_f_month and actual_f_month in filt_fin.columns:
        m_fin = filt_fin.groupby(actual_f_month).agg({
            col_f_total_cost: 'sum', col_f_net_profit: 'sum',
            col_f_sales: 'sum', col_f_overdue_val: 'sum'
        }).fillna(0)
        fig6 = go.Figure()
        for c, n, clr in zip(
            [col_f_total_cost, col_f_net_profit, col_f_sales, col_f_overdue_val],
            ['Cost','Net Profit','Sales','Overdue'],
            ["#126389","#4c8bdc","#2c5aa0","#3ab1dc"]
        ):
            fig6.add_trace(go.Bar(name=n, x=m_fin.index, y=m_fin[c], marker_color=clr))
        fig6.update_layout(height=280, barmode='group', margin=dict(l=20,r=20,t=0,b=30))
        st.plotly_chart(fig6, use_container_width=True, config={'displayModeBar': False})
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="chart-box"><div class="chart-title">Monthly Trends Table</div>', unsafe_allow_html=True)
    if actual_f_month and actual_f_month in filt_fin.columns:
        m_table = filt_fin.groupby(actual_f_month).agg({
            col_f_total_cost: 'sum', col_f_net_profit: 'sum',
            col_f_sales: 'sum', col_f_overdue_val: 'sum'
        }).fillna(0).round(0).astype(int)
        m_table.columns = ['Cost','Profit','Sales','Overdue']
        st.dataframe(m_table, use_container_width=True, height=200)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")
st.markdown(f"<p style='text-align:center; color:#999; font-size:11px;'>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>", unsafe_allow_html=True)
