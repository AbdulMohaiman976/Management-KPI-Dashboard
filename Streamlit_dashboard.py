import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(page_title="Management KPIs Dashboard", page_icon="📊", layout="wide")

# ============ CUSTOM CSS ============
st.markdown("""
<style>
    body { background-color: #f5f5f5; }
    .main { background-color: #f5f5f5; }
    .header-title {
        background: linear-gradient(90deg, #1a3a52 0%, #2c5aa0 100%);
        color: white; padding: 20px; border-radius: 8px; text-align: center;
        margin-bottom: 20px; font-weight: bold; font-size: 28px;
    }
    .filter-box { background-color: white; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
    .kpi-card {
        background-color: white; padding: 20px; border-radius: 6px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1); text-align: center;
        margin-bottom: 12px; border-top: 4px solid #2c5aa0;
        display: flex; flex-direction: column; justify-content: center; align-items: center;
    }
    /* EXACT SAME FONT FOR ALL KPI CARDS - CENTERED */
    .kpi-value { font-size: 28px; font-weight: bold; color: #1a3a52; line-height: 1.2; margin: 10px 0; }
    .kpi-label { font-size: 10px; color: #999; text-transform: uppercase; margin: 0; }
    .chart-box {
        background-color: white; padding: 15px; border-radius: 6px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 15px;
    }
    .chart-title { font-size: 12px; font-weight: bold; color: #1a3a52; margin-bottom: 10px; text-transform: uppercase; }
</style>
""", unsafe_allow_html=True)

# ============ SESSION STATE ============
if 'sales_data' not in st.session_state:
    st.session_state.sales_data = None
if 'financial_data' not in st.session_state:
    st.session_state.financial_data = None
if 'evaluation_data' not in st.session_state:
    st.session_state.evaluation_data = None

# ============ SIDEBAR - FILE UPLOADS ============
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

# ============ CHECK FILES ============
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

# COLUMN NAMES
col_s_year = "Year"
col_s_month = "Month"
col_s_dials = "Dials"
col_s_calls = "Calls"
col_s_dm = "DM Conducted"
col_s_dw = "DW Conducted"
col_s_prop_sent = "Proposals Sent"
col_s_prop_sold = "Proposals Sold"
col_s_sales_val = "Sales Value"
col_s_d2c = "Dials to calls %"
col_s_c2dm = "Calls to DM %"
col_s_dm2dw = "DMs to DW %"
col_s_dw2c = "DWs to Contract %"

col_f_year = "Year"
col_f_month = "Month"
col_f_cashflow = "Cashflow_Coverage_Months"
col_f_sales = "Total_Sales_Value"
col_f_cogs = "Cost_of_Goods_Sold"
col_f_op_cost = "Operating_Cost"
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

# ============ HELPER FUNCTIONS ============
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

# ============ FILTER OPTIONS ============
month_opts = ["All"]
if col_s_month in sales_df.columns:
    months = sorted([str(x) for x in sales_df[col_s_month].dropna().unique()])
    month_opts = ["All"] + months

trainer_opts = ["All"]
if col_e_trainer in evaluation_df.columns:
    trainers = sorted([str(x) for x in evaluation_df[col_e_trainer].dropna().unique()])
    trainer_opts = ["All"] + trainers

# ============ HEADER ============
st.markdown('<div class="header-title">Management KPIs Dashboard</div>', unsafe_allow_html=True)

# ============ FILTERS ============
st.markdown('<div class="filter-box">', unsafe_allow_html=True)
f_col1, f_col2, f_col3 = st.columns(3)
with f_col1: sel_year = st.selectbox("Select Year", ["All"], label_visibility="collapsed")
with f_col2: sel_trainer = st.selectbox("Select Trainer Name", trainer_opts, label_visibility="collapsed")
with f_col3: sel_month = st.selectbox("Select Month", month_opts, label_visibility="collapsed")
st.markdown('</div>', unsafe_allow_html=True)

# ============ FILTER DATA ============
filt_sales = sales_df.copy()
if sel_month != "All" and col_s_month in sales_df.columns:
    filt_sales = filt_sales[filt_sales[col_s_month].astype(str) == str(sel_month)]

filt_eval = evaluation_df.copy()
if sel_trainer != "All" and col_e_trainer in evaluation_df.columns:
    filt_eval = filt_eval[filt_eval[col_e_trainer].astype(str) == str(sel_trainer)]

filt_fin = financial_df.copy()
if sel_month != "All" and col_f_month in financial_df.columns:
    filt_fin = filt_fin[filt_fin[col_f_month].astype(str) == str(sel_month)]

# ============ 3-COLUMN LAYOUT ============
col1, col2, col3 = st.columns(3, gap="small")

# ============ LEFT COLUMN - SALES KPI ============
with col1:
    st.markdown('### Sales KPIs')
    dials = safe_sum(filt_sales, col_s_dials)
    calls = safe_sum(filt_sales, col_s_calls)
    dm = safe_sum(filt_sales, col_s_dm)
    dw = safe_sum(filt_sales, col_s_dw)
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
    
    st.markdown('<div class="chart-box"><div class="chart-title">Summary Chart</div>', unsafe_allow_html=True)
    
    # FIX 1: FUNNEL TEXT SIZE LARGER + WHITE COLOR
    fig1 = go.Figure(go.Funnel(
        y=['Dials', 'Calls', 'DM', 'DW', 'Prop Sent', 'Prop Sold'], 
        x=[dials, calls, dm, dw, prop_sent, prop_sold], 
        textinfo="value+label", 
        textposition="inside",
        insidetextanchor="middle",
        marker=dict(color=['#1a3a52', '#2c5aa0', "#366599", "#266092", "#4b7eb1", "#5e91bd"])
    ))
    
    fig1.update_layout(
        height=300, 
        margin=dict(l=20, r=20, t=0, b=20),
        font=dict(color="white", size=14)
    )
    
    st.plotly_chart(fig1, use_container_width=True, config={'displayModeBar': False})
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="chart-box"><div class="chart-title">Monthly Sales Conversion Trends</div>', unsafe_allow_html=True)
    if col_s_month in filt_sales.columns:
        conv_cols = {}
        if col_s_d2c in filt_sales.columns: conv_cols[col_s_d2c] = lambda x: pd.to_numeric(x, errors='coerce').mean() * 100
        if col_s_c2dm in filt_sales.columns: conv_cols[col_s_c2dm] = lambda x: pd.to_numeric(x, errors='coerce').mean() * 100
        if col_s_dm2dw in filt_sales.columns: conv_cols[col_s_dm2dw] = lambda x: pd.to_numeric(x, errors='coerce').mean() * 100
        if col_s_dw2c in filt_sales.columns: conv_cols[col_s_dw2c] = lambda x: pd.to_numeric(x, errors='coerce').mean() * 100
        if conv_cols:
            monthly_data = filt_sales.groupby(col_s_month).agg(conv_cols).fillna(0)
            # ALL TEXT WHITE
            fig2 = go.Figure()
            for c, name, color in zip([col_s_d2c, col_s_c2dm, col_s_dm2dw, col_s_dw2c], ['Dials→Calls %', 'Calls→DM %', 'DM→DW %', 'DW→Contract %'], ['#1a3a52', '#2c5aa0', '#5a8cc7', '#8bb3d6']):
                if c in monthly_data.columns: fig2.add_trace(go.Bar(name=name, x=monthly_data.index, y=monthly_data[c], marker_color=color, text=[f"{v:.1f}%" for v in monthly_data[c]], textposition='auto', textfont=dict(color='white', size=12)))
            fig2.update_layout(height=280, barmode='group', margin=dict(l=20, r=20, t=0, b=30), yaxis_range=[0, 100], font=dict(color="white", size=11))
            st.plotly_chart(fig2, use_container_width=True, config={'displayModeBar': False})
    st.markdown('</div>', unsafe_allow_html=True)

# ============ MIDDLE COLUMN - EVALUATION ============
with col2:
    st.markdown('### Training Evaluation')
    st.markdown(f'**Trainer: {sel_trainer}**')
    
    st.markdown('<div class="chart-box"><div class="chart-title">Low Expectation over Sessions</div>', unsafe_allow_html=True)
    if col_e_expectation in filt_eval.columns:
        exp_vals = pd.to_numeric(filt_eval[col_e_expectation], errors='coerce').dropna()
        low, good = len(exp_vals[exp_vals <= 3]), len(exp_vals[exp_vals > 3])
        fig3 = go.Figure(data=[go.Pie(labels=['Low (≤3)', 'Good (>3)'], values=[low, good], marker=dict(colors=["#2c78c5", "#4498b9"]), textinfo='label+percent', textfont=dict(color='white', size=12))])
        fig3.update_layout(height=250, margin=dict(l=0, r=0, t=0, b=0), font=dict(color='white'))
        st.plotly_chart(fig3, use_container_width=True, config={'displayModeBar': False})
    st.markdown('</div>', unsafe_allow_html=True)
    
    # FIX 3: TRAINER CHART SIZE LARGER + BETTER X-AXIS ALIGNMENT
    st.markdown('<div class="chart-box"><div class="chart-title">Trainer Performance Comparison</div>', unsafe_allow_html=True)
    perf_cols = {col_e_content: 'Content', col_e_exercise: 'Exercise', col_e_facilitator: 'Facilitator', col_e_expectation: 'Expectation'}
    aggs = {k: (lambda x: pd.to_numeric(x, errors='coerce').mean()) for k in perf_cols.keys() if k in evaluation_df.columns}
    if aggs:
        trainer_perf = evaluation_df.groupby(col_e_trainer).agg(aggs).fillna(0)
        fig4 = go.Figure()
        for k, name, color in zip(aggs.keys(), perf_cols.values(), ['#1a3a52', '#2c5aa0', '#5a8cc7', '#8bb3d6']):
            fig4.add_trace(go.Bar(x=trainer_perf.index, y=trainer_perf[k], name=name, marker_color=color))
        fig4.update_layout(height=320, barmode='group', margin=dict(l=20, r=20, t=0, b=100), xaxis_tickangle=-45, xaxis=dict(tickfont=dict(size=11)))
        st.plotly_chart(fig4, use_container_width=True, config={'displayModeBar': False})
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="chart-box"><div class="chart-title">Month-by-Month Trend Analysis</div>', unsafe_allow_html=True)
    try:
        if col_e_date in evaluation_df.columns:
            eval_df_temp = evaluation_df.copy()
            eval_df_temp['date_dt'] = pd.to_datetime(eval_df_temp[col_e_date], errors='coerce')
            eval_df_temp['month_name'] = eval_df_temp['date_dt'].dt.strftime('%b')
            eval_df_temp['month_num'] = eval_df_temp['date_dt'].dt.month
            
            trend_aggs = {k: (lambda x: pd.to_numeric(x, errors='coerce').mean()) for k in perf_cols.keys() if k in evaluation_df.columns}
            if trend_aggs:
                monthly_trend = eval_df_temp.groupby(['month_num', 'month_name']).agg(trend_aggs).reset_index().sort_values('month_num')
                fig5 = go.Figure()
                for k, name, color in zip(trend_aggs.keys(), perf_cols.values(), ['#1a3a52', '#2c5aa0', '#5a8cc7', '#8bb3d6']):
                    fig5.add_trace(go.Scatter(x=monthly_trend['month_name'], y=monthly_trend[k], mode='lines+markers', name=name, line=dict(color=color, width=3)))
                fig5.update_layout(height=280, margin=dict(l=20, r=20, t=0, b=30), hovermode='x unified')
                st.plotly_chart(fig5, use_container_width=True, config={'displayModeBar': False})
    except: st.info("Trend unavailable")
    st.markdown('</div>', unsafe_allow_html=True)

# ============ RIGHT COLUMN - FINANCIAL KPI ============
with col3:
    st.markdown('### Financial KPIs')
    cf, sl, cg = safe_avg(filt_fin, col_f_cashflow), safe_sum(filt_fin, col_f_sales), safe_sum(filt_fin, col_f_cogs)
    np, oc, ov = safe_sum(filt_fin, col_f_net_profit), safe_sum(filt_fin, col_f_overdue_count), safe_sum(filt_fin, col_f_overdue_val)
    
    kf1, kf2, kf3 = st.columns(3)
    # Format large numbers with K/M suffix to fit in cards
    sl_fmt = f"{int(sl/1000)}K" if sl >= 1000 else str(int(sl))
    cg_fmt = f"{int(cg/1000)}K" if cg >= 1000 else str(int(cg))
    np_fmt = f"{int(np/1000)}K" if np >= 1000 else str(int(np))
    ov_fmt = f"{int(ov/1000)}K" if ov >= 1000 else str(int(ov))
    
    with kf1: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{cf:.1f}</div><div class="kpi-label">Cashflow</div></div>', unsafe_allow_html=True)
    with kf2: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{sl_fmt}</div><div class="kpi-label">Total Sales</div></div>', unsafe_allow_html=True)
    with kf3: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{cg_fmt}</div><div class="kpi-label">Total COGS</div></div>', unsafe_allow_html=True)
    
    kf4, kf5, kf6 = st.columns(3)
    with kf4: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{np_fmt}</div><div class="kpi-label">Net Profit</div></div>', unsafe_allow_html=True)
    with kf5: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{int(oc)}</div><div class="kpi-label">Overdue Count</div></div>', unsafe_allow_html=True)
    with kf6: st.markdown(f'<div class="kpi-card"><div class="kpi-value">{ov_fmt}</div><div class="kpi-label">Overdue Val</div></div>', unsafe_allow_html=True)
    
    st.markdown('<div class="chart-box"><div class="chart-title">Net Profit/Loss Over Month</div>', unsafe_allow_html=True)
    if col_f_month in filt_fin.columns:
        m_fin = filt_fin.groupby(col_f_month).agg({col_f_total_cost: 'sum', col_f_net_profit: 'sum', col_f_sales: 'sum', col_f_overdue_val: 'sum'}).fillna(0)
        fig6 = go.Figure()
        for c, n, clr in zip([col_f_total_cost, col_f_net_profit, col_f_sales, col_f_overdue_val], ['Cost', 'Net Profit', 'Sales', 'Overdue'], ["#126389", "#4c8bdc", '#2c5aa0', "#3ab1dc"]):
            fig6.add_trace(go.Bar(name=n, x=m_fin.index, y=m_fin[c], marker_color=clr))
        fig6.update_layout(height=280, barmode='group', margin=dict(l=20, r=20, t=0, b=30))
        st.plotly_chart(fig6, use_container_width=True, config={'displayModeBar': False})
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="chart-box"><div class="chart-title">Monthly Trends</div>', unsafe_allow_html=True)
    if col_f_month in filt_fin.columns:
        m_table = filt_fin.groupby(col_f_month).agg({col_f_total_cost: 'sum', col_f_net_profit: 'sum', col_f_sales: 'sum', col_f_overdue_val: 'sum'}).fillna(0).round(0).astype(int)
        m_table.columns = ['Cost', 'Profit', 'Sales', 'Overdue']
        st.dataframe(m_table, use_container_width=True, height=200)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")
st.markdown(f"<p style='text-align:center; color:#999; font-size:11px;'>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>", unsafe_allow_html=True)
