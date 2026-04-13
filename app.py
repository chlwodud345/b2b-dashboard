import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import requests

st.set_page_config(page_title="대상웰라이프 B2B몰 대시보드", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

COLORS = ['#3366CC','#E8853D','#27AE60','#9B59B6','#E74C3C','#1ABC9C','#F39C12','#2980B9','#8E44AD','#D35400']
HOVER_FONT = dict(font=dict(size=16, family='Noto Sans KR'))

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700&display=swap');
html, body, [class*="st-"] { font-family: 'Noto Sans KR', sans-serif; }
.main-header { background: linear-gradient(135deg, #1B2A4A 0%, #2D4A7A 100%); color: white; padding: 24px 32px; border-radius: 14px; margin-bottom: 28px; }
.main-header h1 { margin: 0; font-size: 1.6rem; font-weight: 700; }
.main-header p { margin: 6px 0 0; opacity: 0.7; font-size: 0.9rem; }
.kpi-card { background: white; border-radius: 14px; padding: 24px; border: 1px solid #e2e8f0; box-shadow: 0 2px 6px rgba(0,0,0,0.04); text-align: center; margin-bottom: 20px; }
.kpi-label { font-size: 0.85rem; color: #64748b; font-weight: 500; margin-bottom: 6px; }
.kpi-value { font-size: 1.8rem; font-weight: 700; color: #1e293b; }
.kpi-unit { font-size: 0.9rem; color: #94a3b8; margin-left: 3px; }
.stTabs [data-baseweb="tab-list"] { gap: 4px; }
.stTabs [data-baseweb="tab"] { padding: 12px 24px; font-weight: 500; font-size: 0.95rem; }
[data-testid="stSidebar"] { background: #f8fafc; }
.stDataFrame [data-testid="column-header-menu"] { z-index: 999 !important; background: white !important; box-shadow: 0 4px 12px rgba(0,0,0,0.15) !important; }
.js-plotly-plot .plotly .hoverlayer .hovertext text { font-size: 22px !important; }
.js-plotly-plot .plotly .hoverlayer .hovertext path { stroke-width: 2px !important; }
div[data-testid="stVerticalBlock"] > div { padding-top: 4px; padding-bottom: 4px; }
</style>
""", unsafe_allow_html=True)

def fmt_krw(n):
    if pd.isna(n) or n == 0: return "0원"
    sign = '' if n >= 0 else '-'
    a = abs(n)
    if a >= 1e8: return f"{sign}{a/1e8:.1f}억원"
    if a >= 1e4: return f"{sign}{a/1e4:,.0f}만원"
    return f"{sign}{a:,.0f}원"
def fmt_krw_short(n):
    if pd.isna(n) or n == 0: return "0"
    sign = '' if n >= 0 else '-'
    a = abs(n)
    if a >= 1e8: return f"{sign}{a/1e8:.1f}억"
    if a >= 1e4: return f"{sign}{a/1e4:,.0f}만"
    return f"{sign}{a:,.0f}"
def fmt_num(n):
    if pd.isna(n): return "0"
    return f"{n:,.0f}"
def fmt_pct(n):
    if pd.isna(n): return "0%"
    return f"{n:.1f}%"
def kpi_card(label, value, unit=""):
    return f'<div class="kpi-card"><div class="kpi-label">{label}</div><div class="kpi-value">{value}<span class="kpi-unit">{unit}</span></div></div>'
def to_ym_kr(ym_str):
    if not ym_str or pd.isna(ym_str): return ''
    parts = str(ym_str).split('-')
    if len(parts) >= 2: return f"{parts[0]}년 {int(parts[1])}월"
    return str(ym_str)
def ym_series_kr(series):
    return series.apply(to_ym_kr)
def to_date_kr(date_str):
    if not date_str or pd.isna(date_str): return ''
    parts = str(date_str).split('-')
    if len(parts) == 3: return f"{parts[0]}년 {int(parts[1])}월 {int(parts[2])}일"
    return str(date_str)
def krw_tickvals(series, n=5):
    mn, mx = series.min(), series.max()
    if mx == 0: return [0], ['0']
    vals = np.linspace(0, mx * 1.05, n).tolist()
    texts = [fmt_krw_short(v) for v in vals]
    return vals, texts
def make_donut(df, name_col, value_col, colors=None):
    total = df[value_col].sum()
    fig = go.Figure()
    fig.add_trace(go.Pie(labels=df[name_col], values=df[value_col], hole=0.55,
        marker=dict(colors=(colors or COLORS)[:len(df)]),
        textinfo='label+percent', textposition='inside', insidetextorientation='horizontal', textfont=dict(size=11),
        hovertemplate='%{label}<br>매출: %{customdata}<br>비중: %{percent}<extra></extra>',
        customdata=[f"{v:,.0f}원" for v in df[value_col]]))
    fig.add_annotation(text=f"<b>합계</b><br>{fmt_krw(total)}", x=0.5, y=0.5, font=dict(size=15, color='#1e293b'), showarrow=False, xref='paper', yref='paper')
    fig.update_layout(height=520, margin=dict(l=20, r=20, t=30, b=140),
        legend=dict(orientation="h", yanchor="top", y=-0.02, xanchor="center", x=0.5, font=dict(size=11), traceorder='normal'), showlegend=True)
    return fig

# ============================================================
# 데이터 전처리
# ============================================================
@st.cache_data
def process_data(orders, members, referrals_df):
    orders['주문일'] = pd.to_datetime(orders['주문일'], errors='coerce')
    orders['주문일자'] = orders['주문일'].dt.strftime('%Y-%m-%d')
    orders['주문월'] = orders['주문일'].dt.to_period('M').astype(str)
    orders['주문시간'] = orders['주문일'].dt.hour
    dow_map = {'Monday':'월','Tuesday':'화','Wednesday':'수','Thursday':'목','Friday':'금','Saturday':'토','Sunday':'일'}
    orders['요일'] = orders['주문일'].dt.day_name().map(dow_map)
    def parse_num(val):
        if pd.isna(val) or val == '-': return 0
        if isinstance(val, (int, float)): return val
        try: return float(str(val).replace(',','').replace('원','').strip())
        except: return 0
    orders['판매합계금액'] = orders['판매합계금액'].apply(parse_num)
    orders['주문 수량'] = orders['주문 수량'].apply(parse_num)
    def extract_region(addr):
        if pd.isna(addr): return '기타'
        addr = str(addr).strip()
        for r in ['서울','부산','대구','인천','광주','대전','울산','세종','경기','강원','충북','충남','전북','전남','경북','경남','제주']:
            if addr.startswith(r): return r
        return '기타'
    orders['지역'] = orders['주소(주문자)'].apply(extract_region)
    members['가입일'] = pd.to_datetime(members['가입일'], errors='coerce')
    members['가입월'] = members['가입일'].dt.to_period('M').astype(str)
    members['사업자번호'] = members['사업자번호'].astype(str).str.replace('-','').str.strip()
    referrals_df['피추천인 사업자 번호'] = referrals_df['피추천인 사업자 번호'].astype(str).str.replace('-','').str.strip()
    return orders, members, referrals_df

@st.cache_data
def process_bw(bw_raw):
    bw = bw_raw.copy()
    # 컬럼명 정리 (줄바꿈 제거)
    bw.columns = [c.replace('\n','').strip() for c in bw.columns]
    # 달력 연도/월 → 문자열 YYYY-MM
    def parse_ym(val):
        if pd.isna(val): return None
        s = str(val).strip()
        if '.' in s:
            parts = s.split('.')
            y = parts[0]
            m = parts[1].zfill(2)
            return f"{y}-{m}"
        return s
    bw['연월'] = bw['달력 연도/월'].apply(parse_ym)
    bw['연도'] = bw['연월'].str[:4]
    bw['월'] = bw['연월'].str[5:7].astype(int, errors='ignore')
    # 고객명 → 채널 라벨 매핑
    def customer_label(name):
        s = str(name).strip()
        parts = s.split(',')
        if len(parts) == 1: return '일반'
        ref_map = {'영':'영업','대':'대리점','케':'케어포'}
        mem_map = {'의':'의료기','장':'장기요양','병':'병원','약':'약국','크':'염증성장질환','종':'종사자'}
        if len(parts) == 2:
            return ref_map.get(parts[1].strip(), parts[1].strip())
        if len(parts) == 3:
            p1 = parts[1].strip()
            p2 = parts[2].strip()
            if p2 == '종':
                prefix = mem_map.get(p1, ref_map.get(p1, p1))
                return f"{prefix}-종사자"
            r = ref_map.get(p1, p1)
            m = mem_map.get(p2, p2)
            return f"{r}-{m}"
        return s
    bw['채널'] = bw['고객명'].apply(customer_label)
    # 기타판관비 계산
    bw['기타판관비'] = bw['IV.판매비 및 관리비'] - bw['IV.6.광고선전비'] - bw['IV.7.운반비'] - bw['IV.8.판매수수료'] - bw['IV.9.판촉비']
    # 이익률 계산
    bw['매출총이익률'] = np.where(bw['I.매출액(FI기준)'] != 0, bw['III.매출총이익'] / bw['I.매출액(FI기준)'] * 100, 0)
    bw['영업이익률'] = np.where(bw['I.매출액(FI기준)'] != 0, bw['V.영업이익I'] / bw['I.매출액(FI기준)'] * 100, 0)
    return bw

# ============================================================
# 데이터 로드 (구글 드라이브)
# ============================================================
GDRIVE_FILE_ID = '1Op9Y2FFb_aLQJKAcLyKj9HJQbK6YYnmf'
def download_from_gdrive(file_id):
    import gdown
    import tempfile, os
    tmp = tempfile.mktemp(suffix='.xlsx')
    url = f'https://drive.google.com/uc?id={file_id}'
    gdown.download(url, tmp, quiet=True)
    with open(tmp, 'rb') as f:
        content = f.read()
    os.remove(tmp)
    return io.BytesIO(content)

@st.cache_data(ttl=3600, show_spinner="📥 구글 드라이브에서 데이터를 불러오는 중...")
def load_from_gdrive():
    fb = download_from_gdrive(GDRIVE_FILE_ID)
    o = pd.read_excel(fb, sheet_name='주문내역', header=1, engine='openpyxl'); fb.seek(0)
    m = pd.read_excel(fb, sheet_name='회원정보', header=1, engine='openpyxl'); fb.seek(0)
    r = pd.read_excel(fb, sheet_name='추천인', header=1, engine='openpyxl'); fb.seek(0)
    # BW 시트 로드
    try:
        bw_raw = pd.read_excel(fb, sheet_name='BW', header=0, engine='openpyxl', 
                        dtype={'달력 연도/월': str})
        bw = process_bw(bw_raw)
    except Exception:
        bw = pd.DataFrame()
    orders, members, referrals_df = process_data(o, m, r)
    return orders, members, referrals_df, bw

try:
    orders, members, referrals_df, bw_data = load_from_gdrive()
    sidebar_msg = f"✅ 데이터 로드 완료\n- 주문: {len(orders):,}건\n- 회원: {len(members):,}건\n- 추천인: {len(referrals_df):,}건"
    if len(bw_data) > 0:
        sidebar_msg += f"\n- BW손익: {len(bw_data):,}건"
    st.sidebar.success(sidebar_msg)
except Exception as e:
    st.error(f"❌ 데이터 로드 실패: {str(e)}\n\n구글 드라이브 파일 공유 설정을 확인해주세요.")
    st.stop()
if st.sidebar.button("🔄 데이터 새로고침"):
    st.cache_data.clear()
    st.rerun()
st.sidebar.markdown("---")
member_lookup = members.set_index('아이디')[['상호명','회원타입','회원등급']].to_dict('index')

st.sidebar.markdown("## 🔍 필터")
years = sorted(orders['주문일'].dt.year.dropna().unique().astype(int))
selected_years = st.sidebar.multiselect("연도", [str(y) for y in years], default=[], placeholder="전체")
if selected_years:
    all_months = set()
    for y in selected_years:
        ms = orders[orders['주문일'].dt.year == int(y)]['주문일'].dt.month.dropna().unique().astype(int)
        all_months.update(ms)
    selected_months = st.sidebar.multiselect("월", [f"{m}월" for m in sorted(all_months)], default=[], placeholder="전체")
else:
    selected_months = st.sidebar.multiselect("월", [f"{m}월" for m in range(1,13)], default=[], placeholder="전체")
type_opts = sorted(orders['주문자 구분'].dropna().unique().tolist())
selected_types = st.sidebar.multiselect("회원구분", type_opts, default=[], placeholder="전체")
grade_opts = sorted(orders['회원 등급'].dropna().unique().tolist())
selected_grades = st.sidebar.multiselect("회원등급", grade_opts, default=[], placeholder="전체")

filtered = orders.copy()
if selected_years: filtered = filtered[filtered['주문일'].dt.year.isin([int(y) for y in selected_years])]
if selected_months: filtered = filtered[filtered['주문일'].dt.month.isin([int(m.replace('월','')) for m in selected_months])]
if selected_types: filtered = filtered[filtered['주문자 구분'].isin(selected_types)]
if selected_grades: filtered = filtered[filtered['회원 등급'].isin(selected_grades)]
filtered_members = members.copy()
if selected_years: filtered_members = filtered_members[filtered_members['가입일'].dt.year.isin([int(y) for y in selected_years])]
if selected_months: filtered_members = filtered_members[filtered_members['가입일'].dt.month.isin([int(m.replace('월','')) for m in selected_months])]
if selected_types: filtered_members = filtered_members[filtered_members['회원타입'].isin(selected_types)]
if selected_grades: filtered_members = filtered_members[filtered_members['회원등급'].isin(selected_grades)]

# BW 필터 (연도/월만 적용)
bw_filtered = bw_data.copy()
if len(bw_filtered) > 0:
    if selected_years: bw_filtered = bw_filtered[bw_filtered['연도'].isin(selected_years)]
    if selected_months: bw_filtered = bw_filtered[bw_filtered['월'].isin([int(m.replace('월','')) for m in selected_months])]

import base64, os
logo_path = os.path.join(os.path.dirname(__file__), 'logo.png')
if os.path.exists(logo_path):
    with open(logo_path, 'rb') as f:
        logo_b64 = base64.b64encode(f.read()).decode()
    st.markdown(f'<div class="main-header" style="display:flex;align-items:center;justify-content:space-between;"><div><h1>📊 대상웰라이프 B2B몰 대시보드</h1><p>Sales & Operations Analytics</p></div><img src="data:image/png;base64,{logo_b64}" style="height:50px;object-fit:contain;"></div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="main-header"><h1>📊 대상웰라이프 B2B몰 대시보드</h1><p>Sales & Operations Analytics</p></div>', unsafe_allow_html=True)

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(["📋 종합 현황","💰 매출 분석","📦 상품 분석","👥 회원 분석","🔗 추천인 분석","💚 케어포 멤버십","📈 손익 분석"])

# ============================================================
# Tab 1. 종합 현황
# ============================================================
with tab1:
    ts = filtered['판매합계금액'].sum(); to_ = filtered['주문 ID'].nunique(); tb = filtered['주문자 ID'].nunique()
    tm = len(members); nm = len(filtered_members); ao = ts/to_ if to_>0 else 0
    cols = st.columns(6)
    for col,(l,v,u) in zip(cols,[("총 매출액",fmt_krw_short(ts),"원"),("총 주문건수",fmt_num(to_),"건"),("총 회원수",fmt_num(tm),"처"),("주문회원수",fmt_num(tb),"처"),("신규 가입회원",fmt_num(nm),"처"),("객단가",fmt_krw_short(ao),"원")]):
        col.markdown(kpi_card(l,v,u), unsafe_allow_html=True)
    st.markdown("#### 월별 매출 · 주문건수 추이")
    monthly = filtered.groupby('주문월').agg(매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique')).reset_index()
    monthly['주문월_kr'] = ym_series_kr(monthly['주문월'])
    fig = make_subplots(specs=[[{"secondary_y":True}]])
    fig.add_trace(go.Bar(x=monthly['주문월_kr'],y=monthly['매출'],name='매출액',marker_color='#3366CC',opacity=0.8,text=[fmt_krw_short(v) for v in monthly['매출']],textposition='outside',textfont=dict(size=11),hovertemplate='%{x}<br>매출: %{customdata}<extra></extra>',customdata=[f"{v:,.0f}원" for v in monthly['매출']]),secondary_y=False)
    fig.add_trace(go.Scatter(x=monthly['주문월_kr'],y=monthly['주문건수'],name='주문건수',line=dict(color='#E8853D',width=3),mode='lines+markers',marker=dict(size=8),hovertemplate='%{x}<br>주문: %{y:,}건<extra></extra>'),secondary_y=True)
    tvals, ttexts = krw_tickvals(monthly['매출'])
    fig.update_layout(height=480,margin=dict(l=80,r=60,t=50,b=70),legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=12)),xaxis=dict(tickfont=dict(size=12)))
    fig.update_yaxes(title_text="매출액",tickvals=tvals,ticktext=ttexts,tickfont=dict(size=11),secondary_y=False)
    fig.update_yaxes(title_text="주문건수",tickfont=dict(size=11),secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)
    cl,cr = st.columns(2)
    with cl:
        st.markdown("#### 회원구분별 매출 비중")
        ts_df = filtered.groupby('주문자 구분')['판매합계금액'].sum().reset_index(); ts_df.columns = ['구분','매출']; ts_df = ts_df.sort_values('매출',ascending=False)
        fig = make_donut(ts_df, '구분', '매출'); fig.update_layout(height=520)
        st.plotly_chart(fig, use_container_width=True)
    with cr:
        st.markdown("#### 지역별 매출")
        rg = filtered.groupby('지역')['판매합계금액'].sum().sort_values().reset_index(); rg.columns = ['지역','매출']
        fig = px.bar(rg,x='매출',y='지역',orientation='h',color_discrete_sequence=COLORS)
        fig.update_traces(text=[fmt_krw_short(v) for v in rg['매출']],textposition='outside',textfont=dict(size=11),hovertemplate='%{y}: %{customdata}<extra></extra>',customdata=[fmt_krw(v) for v in rg['매출']])
        fig.update_layout(height=520,margin=dict(l=70,r=100,t=30,b=40),showlegend=False,xaxis=dict(title='',tickfont=dict(size=11)),yaxis=dict(title='',tickfont=dict(size=12)))
        st.plotly_chart(fig, use_container_width=True)
    st.markdown("#### 일별 매출 추이")
    daily = filtered.groupby('주문일자')['판매합계금액'].sum().reset_index(); daily.columns = ['날짜','매출']
    fig = px.area(daily,x='날짜',y='매출',color_discrete_sequence=['#3366CC'])
    fig.update_traces(customdata=list(zip(daily['날짜'].apply(to_date_kr),[fmt_krw(v) for v in daily['매출']])),hovertemplate='%{customdata[0]}<br>매출: %{customdata[1]}<extra></extra>')
    tvals2, ttexts2 = krw_tickvals(daily['매출'])
    fig.update_layout(height=400,margin=dict(l=80,r=30,t=30,b=60),showlegend=False,xaxis=dict(title='날짜',tickfont=dict(size=11),title_font=dict(size=13),tickformat='%Y년 %m월'),yaxis=dict(title='매출액',tickvals=tvals2,ticktext=ttexts2,tickfont=dict(size=11),title_font=dict(size=13)))
    st.plotly_chart(fig, use_container_width=True)

# ============================================================
# Tab 2. 매출 분석
# ============================================================
with tab2:
    st.markdown("#### 회원구분별 × 월별 매출 추이")
    tm_df = filtered.groupby(['주문월','주문자 구분'])['판매합계금액'].sum().reset_index(); tm_df['주문월_kr'] = ym_series_kr(tm_df['주문월'])
    fig = px.bar(tm_df,x='주문월_kr',y='판매합계금액',color='주문자 구분',color_discrete_sequence=COLORS)
    for tr in fig.data: tr.customdata = [fmt_krw(v) for v in tr.y]; tr.hovertemplate = '%{x}<br>' + tr.name + ': %{customdata}<extra></extra>'
    fig.update_layout(height=480,barmode='stack',margin=dict(l=70,r=30,t=50,b=70),legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=11)),xaxis=dict(title='',tickfont=dict(size=12)),yaxis=dict(title='매출액',tickfont=dict(size=11)))
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("#### 회원등급별 매출")
    gs = filtered.groupby('회원 등급').agg(매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique'),주문회원수=('주문자 ID','nunique')).reset_index().sort_values('매출')
    fig = px.bar(gs,x='매출',y='회원 등급',orientation='h',color_discrete_sequence=COLORS)
    fig.update_traces(text=[fmt_krw_short(v) for v in gs['매출']],textposition='outside',textfont=dict(size=11),hovertemplate='%{y}<br>매출: %{customdata[0]}<br>주문: %{customdata[1]:,}건<br>회원: %{customdata[2]:,}처<extra></extra>',customdata=list(zip([fmt_krw(v) for v in gs['매출']], gs['주문건수'], gs['주문회원수'])))
    fig.update_layout(height=max(420,len(gs)*35+140),margin=dict(l=140,r=100,t=30,b=40),showlegend=False,xaxis=dict(title='매출액',tickfont=dict(size=11)),yaxis=dict(title='',tickfont=dict(size=12)))
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("#### 요일 · 시간대별 주문 매출 히트맵")
    dow_order = ['월','화','수','목','금','토','일']
    hm = filtered.groupby(['요일','주문시간'])['판매합계금액'].sum().reset_index()
    hmp = hm.pivot_table(index='요일',columns='주문시간',values='판매합계금액',fill_value=0).reindex(dow_order)
    fig = go.Figure(data=go.Heatmap(z=hmp.values,x=[f'{h}시' for h in hmp.columns],y=hmp.index,colorscale=[[0,'#F0F2F5'],[0.5,'#6B9BD2'],[1,'#1B2A4A']],text=[[fmt_krw_short(v) for v in row] for row in hmp.values],texttemplate='%{text}',textfont=dict(size=12),hovertemplate='%{y} %{x}<br>매출: %{customdata}<extra></extra>',customdata=[[f"{v:,.0f}원" for v in row] for row in hmp.values]))
    fig.update_layout(height=320,margin=dict(l=50,r=20,t=20,b=40),xaxis=dict(tickfont=dict(size=11)),yaxis=dict(tickfont=dict(size=12),autorange='reversed'))
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("#### 기관별 매출 현황")
    ba = filtered.groupby(['주문자 ID','주문자명','주문자 구분','회원 등급']).agg(매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique'),최근주문일=('주문일자','max')).reset_index()
    ba['객단가'] = (ba['매출']/ba['주문건수']).round(0); ba['상호명'] = ba['주문자 ID'].map(lambda x: member_lookup.get(x, {}).get('상호명', ''))
    ba = ba[['주문자 ID','주문자명','상호명','주문자 구분','회원 등급','주문건수','매출','객단가','최근주문일']].sort_values('매출',ascending=False)
    search = st.text_input("🔍 검색 (아이디, 주문자명, 상호명)",key="sales_search")
    if search: ba = ba[ba.apply(lambda r:search.lower() in str(r).lower(),axis=1)]
    st.dataframe(ba.style.format({'매출':'{:,.0f}원','주문건수':'{:,.0f}건','객단가':'{:,.0f}원'}),use_container_width=True,height=550)

# ============================================================
# Tab 3. 상품 분석
# ============================================================
with tab3:
    pa = filtered.groupby(['상품명','상품 코드']).agg(매출=('판매합계금액','sum'),수량=('주문 수량','sum'),주문건수=('주문 ID','nunique')).reset_index().sort_values('매출',ascending=False)
    st.markdown("#### 상품별 매출 TOP 20 (파레토 차트)")
    top20 = pa.head(20).copy(); ttl = pa['매출'].sum(); top20['누적비중'] = (top20['매출'].cumsum()/ttl*100).round(1)
    fig = make_subplots(specs=[[{"secondary_y":True}]])
    fig.add_trace(go.Bar(x=[f"{i+1}. {n[:16]}" for i,n in enumerate(top20['상품명'])],y=top20['매출'],name='매출액',marker_color='#3366CC',opacity=0.8,hovertemplate='%{customdata[0]}<br>매출: %{customdata[1]}<extra></extra>',customdata=list(zip(top20['상품명'],[fmt_krw(v) for v in top20['매출']]))),secondary_y=False)
    fig.add_trace(go.Scatter(x=[f"{i+1}. {n[:16]}" for i,n in enumerate(top20['상품명'])],y=top20['누적비중'],name='누적 비중',line=dict(color='#E8853D',width=3),mode='lines+markers',marker=dict(size=7),hovertemplate='%{customdata}<br>누적 비중: %{y:.1f}%<extra></extra>',customdata=top20['상품명'].tolist()),secondary_y=True)
    tvals3, ttexts3 = krw_tickvals(top20['매출'])
    fig.update_layout(height=540,margin=dict(l=80,r=60,t=50,b=150),legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=12)),xaxis=dict(tickangle=45,tickfont=dict(size=9)))
    fig.update_yaxes(title_text="매출액",tickvals=tvals3,ticktext=ttexts3,tickfont=dict(size=11),secondary_y=False)
    fig.update_yaxes(title_text="누적 비중 (%)",range=[0,100],ticksuffix='%',tickfont=dict(size=11),secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("#### 전체 상품 매출 현황")
    sp = st.text_input("🔍 상품명/코드 검색",key="product_search"); dp = pa.copy()
    if sp: dp = dp[dp.apply(lambda r:sp.lower() in str(r).lower(),axis=1)]
    st.dataframe(dp.style.format({'매출':'{:,.0f}원','수량':'{:,.0f}','주문건수':'{:,.0f}건'}),use_container_width=True,height=450)
    st.markdown("#### 회원구분별 × 상품 매출 크로스 (TOP 20)")
    t20n = pa.head(20)['상품명'].tolist()
    cp = filtered[filtered['상품명'].isin(t20n)].pivot_table(index='상품명',columns='주문자 구분',values='판매합계금액',aggfunc='sum',fill_value=0)
    cp['합계'] = cp.sum(axis=1); cp = cp.sort_values('합계',ascending=False)
    st.dataframe(cp.style.format('{:,.0f}원'),use_container_width=True,height=550)
    st.markdown("#### 월별 상품 매출 추이")
    t5 = pa.head(5)['상품명'].tolist(); sel = st.multiselect("상품 선택",pa['상품명'].tolist(),default=t5,key="product_trend")
    if sel:
        td = filtered[filtered['상품명'].isin(sel)].groupby(['주문월','상품명'])['판매합계금액'].sum().reset_index(); td['주문월_kr'] = ym_series_kr(td['주문월'])
        fig = px.line(td,x='주문월_kr',y='판매합계금액',color='상품명',markers=True,color_discrete_sequence=COLORS)
        for tr in fig.data:
            tr.customdata = [fmt_krw(v) for v in tr.y]; tr.hovertemplate = '%{x}<br>' + (tr.name[:20]+'...' if len(tr.name)>20 else tr.name) + '<br>매출: %{customdata}<extra></extra>'
            if len(tr.name) > 22: tr.name = tr.name[:22]+'...'
        fig.update_layout(height=480,margin=dict(l=70,r=30,t=30,b=120),legend=dict(orientation="h",yanchor="top",y=-0.15,x=0,font=dict(size=10)),xaxis=dict(title='',tickfont=dict(size=12)),yaxis=dict(title='매출액',tickfont=dict(size=11)))
        st.plotly_chart(fig, use_container_width=True)

# ============================================================
# Tab 4. 회원 분석
# ============================================================
with tab4:
    mo_df = orders.groupby('주문자 ID').agg(첫주문일=('주문일','min'),주문건수=('주문 ID','nunique'),주문월수=('주문월','nunique')).reset_index()
    conv = members[members['아이디'].isin(orders['주문자 ID'].unique())]; conv_r = len(conv)/len(members)*100 if len(members)>0 else 0
    rep = mo_df[mo_df['주문건수']>=2]; rep_r = len(rep)/len(mo_df)*100 if len(mo_df)>0 else 0
    r3m = orders[orders['주문일']>=orders['주문일'].max()-pd.DateOffset(months=3)]; act = r3m['주문자 ID'].nunique()
    cols = st.columns(5)
    for col,(l,v,u) in zip(cols,[("총 회원수",fmt_num(len(members)),"처"),("신규 가입회원",fmt_num(len(filtered_members)),"처"),("구매전환율",fmt_pct(conv_r),""),("재구매율",fmt_pct(rep_r),""),("활성회원(3개월)",fmt_num(act),"처")]):
        col.markdown(kpi_card(l,v,u),unsafe_allow_html=True)
    st.markdown("#### 🔍 회원 상세 검색")
    mem_search = st.text_input("상호명, 아이디, 담당자명으로 검색", key="member_detail_search", placeholder="예: 대상병원, 대상요양원 등")
    if mem_search:
        mask = members.apply(lambda r: mem_search.lower() in str(r.get('상호명','')).lower() or mem_search.lower() in str(r.get('아이디','')).lower() or mem_search.lower() in str(r.get('담당자 이름','')).lower(), axis=1)
        search_members = members[mask].copy()
        if len(search_members) == 0:
            st.warning("검색 결과가 없습니다.")
        else:
            order_agg = orders.groupby('주문자 ID').agg(총매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique'),첫주문일=('주문일자','min'),최근주문일=('주문일자','max')).reset_index()
            order_agg['객단가'] = (order_agg['총매출']/order_agg['주문건수']).round(0)
            top_products = orders.groupby(['주문자 ID','상품명'])['판매합계금액'].sum().reset_index().sort_values(['주문자 ID','판매합계금액'],ascending=[True,False])
            top_products = top_products.groupby('주문자 ID').head(3).groupby('주문자 ID')['상품명'].apply(lambda x: ' / '.join(x)).reset_index(); top_products.columns = ['주문자 ID','주요 구매상품']
            ref_info = referrals_df.groupby('피추천인 사업자 번호').first()[['추천인','회원그룹']].reset_index(); ref_info.columns = ['사업자번호','추천인명','추천인유형']
            result = search_members[['아이디','상호명','사업자번호','회원타입','회원등급','가입일','담당자 이름','휴대폰','주소']].copy()
            result['가입일'] = result['가입일'].dt.strftime('%Y-%m-%d')
            result = result.merge(order_agg, left_on='아이디', right_on='주문자 ID', how='left').drop(columns=['주문자 ID'],errors='ignore')
            result = result.merge(top_products, left_on='아이디', right_on='주문자 ID', how='left').drop(columns=['주문자 ID'],errors='ignore')
            result = result.merge(ref_info, on='사업자번호', how='left')
            for c in ['총매출','객단가']: result[c] = result[c].fillna(0)
            result['주문건수'] = result['주문건수'].fillna(0).astype(int)
            for c in ['주요 구매상품','추천인명','추천인유형']: result[c] = result[c].fillna('-')
            display_cols = ['아이디','상호명','담당자 이름','회원타입','회원등급','가입일','총매출','주문건수','객단가','첫주문일','최근주문일','주요 구매상품','추천인명','추천인유형','휴대폰','주소']
            result = result[[c for c in display_cols if c in result.columns]].sort_values('총매출',ascending=False)
            st.markdown(f"**검색 결과: {len(result)}건**")
            st.dataframe(result.style.format({'총매출':'{:,.0f}원','주문건수':'{:,.0f}건','객단가':'{:,.0f}원'}),use_container_width=True, height=400)
    st.markdown("#### 월별 신규가입자 추이 (회원타입별)")
    jm = filtered_members.groupby(['가입월','회원타입']).size().reset_index(name='가입자수'); jm['가입월_kr'] = ym_series_kr(jm['가입월'])
    fig = px.bar(jm,x='가입월_kr',y='가입자수',color='회원타입',color_discrete_sequence=COLORS)
    for tr in fig.data: tr.hovertemplate = '%{x}<br>' + tr.name + ': %{y:,}처<extra></extra>'
    fig.update_layout(height=480,barmode='stack',margin=dict(l=60,r=30,t=50,b=70),legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=11)),xaxis=dict(title='',tickfont=dict(size=12)),yaxis=dict(title='가입자 수 (처)',tickfont=dict(size=11)))
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("#### 회원등급별 가입자 분포")
    gd = filtered_members['회원등급'].value_counts().reset_index(); gd.columns = ['등급','수']
    fig = px.bar(gd,x='수',y='등급',orientation='h',color_discrete_sequence=COLORS)
    fig.update_traces(text=[fmt_num(v) for v in gd['수']],textposition='outside',textfont=dict(size=11),hovertemplate='%{y}: %{x:,}처<extra></extra>')
    fig.update_layout(height=max(420,len(gd)*35+140),margin=dict(l=140,r=80,t=30,b=40),showlegend=False,xaxis=dict(title='가입자 수 (처)',tickfont=dict(size=11)),yaxis=dict(title='',tickfont=dict(size=12)))
    st.plotly_chart(fig, use_container_width=True)
    cl,cr = st.columns(2)
    with cl:
        st.markdown("#### 가입 후 첫 주문까지 소요일")
        mg = filtered_members.merge(mo_df,left_on='아이디',right_on='주문자 ID',how='inner'); mg['소요일'] = (mg['첫주문일']-mg['가입일']).dt.days; mg = mg[mg['소요일']>=0]
        bins=[0,1,8,15,31,61,91,9999]; lb=['당일','1~7일','8~14일','15~30일','31~60일','61~90일','90일+']
        mg['구간'] = pd.cut(mg['소요일'],bins=bins,labels=lb,right=False); dh = mg['구간'].value_counts().reindex(lb).fillna(0).reset_index(); dh.columns=['구간','회원수']
        fig = px.bar(dh,x='구간',y='회원수',color_discrete_sequence=['#3366CC'])
        fig.update_traces(text=[fmt_num(v) for v in dh['회원수']],textposition='outside',textfont=dict(size=11),hovertemplate='%{x}: %{y:,}처<extra></extra>')
        fig.update_layout(height=450,margin=dict(l=60,r=30,t=30,b=60),showlegend=False,xaxis=dict(title='',tickfont=dict(size=12)),yaxis=dict(title='회원 수 (처)',tickfont=dict(size=11)))
        st.plotly_chart(fig, use_container_width=True)
    with cr:
        st.markdown("#### 주문횟수 구간별 회원 분포")
        mo_df2 = mo_df.copy(); bo=[1,2,4,6,11,21,9999]; lo=['1회','2~3회','4~5회','6~10회','11~20회','20회+']
        mo_df2['구간'] = pd.cut(mo_df2['주문건수'],bins=bo,labels=lo,right=False); od = mo_df2['구간'].value_counts().reindex(lo).fillna(0).reset_index(); od.columns=['구간','회원수']
        fig = px.bar(od,x='구간',y='회원수',color_discrete_sequence=COLORS)
        fig.update_traces(text=[fmt_num(v) for v in od['회원수']],textposition='outside',textfont=dict(size=11),hovertemplate='%{x}: %{y:,}처<extra></extra>')
        fig.update_layout(height=450,margin=dict(l=60,r=30,t=30,b=60),showlegend=False,xaxis=dict(title='',tickfont=dict(size=12)),yaxis=dict(title='회원 수 (처)',tickfont=dict(size=11)))
        st.plotly_chart(fig, use_container_width=True)
    st.markdown("#### 코호트 리텐션 히트맵 (가입월 × 경과월별 재구매율)")
    cm = members[members['가입일'].notna()].copy(); cm['코호트'] = cm['가입일'].dt.to_period('M').astype(str)
    om = orders.groupby('주문자 ID')['주문월'].apply(set).to_dict(); cohorts = sorted(cm['코호트'].unique()); mx=12; rd=[]
    for c in cohorts:
        cu = cm[cm['코호트']==c]['아이디'].tolist(); sz=len(cu)
        if sz==0: continue
        row={'코호트':c,'크기':sz}; cp=pd.Period(c,freq='M')
        for o in range(mx):
            t=str(cp+o); a=sum(1 for u in cu if t in om.get(u,set())); row[f'{o}개월']=round(a/sz*100,1)
        rd.append(row)
    if rd:
        rdf=pd.DataFrame(rd); mc=[f'{o}개월' for o in range(mx)]; zv=rdf[mc].values
        fig=go.Figure(data=go.Heatmap(z=zv,x=mc,y=[f"{to_ym_kr(r['코호트'])} ({r['크기']}처)" for _,r in rdf.iterrows()],colorscale=[[0,'#F0F2F5'],[0.3,'#A8D5A2'],[1,'#27AE60']],text=[[f'{v:.1f}%' if v>0 else '-' for v in row] for row in zv],texttemplate='%{text}',textfont=dict(size=10),hovertemplate='%{y}<br>%{x}: %{z:.1f}%<extra></extra>'))
        fig.update_layout(height=max(400,len(rd)*32+140),margin=dict(l=190,r=20,t=20,b=50),yaxis=dict(tickfont=dict(size=10),autorange="reversed"),xaxis=dict(tickfont=dict(size=11)))
        st.plotly_chart(fig, use_container_width=True)

# ============================================================
# Tab 5. 추천인 분석
# ============================================================
with tab5:
    clm={}
    for _,r in referrals_df.iterrows():
        n=str(r.get('추천인','')).strip(); g=str(r.get('회원그룹',''))
        if not n or n in ['-','nan']: continue
        if g=='영업팀': clm[n]='케어포' if n=='케어포' else '영업팀'
        elif g=='대리점 회원': clm[n]='대리점'
    ra={}
    for _,r in referrals_df.iterrows():
        n=str(r.get('추천인','')).strip()
        if not n or n in ['-','nan']: continue
        if n not in ra: ra[n]={'추천인':n,'유형':clm.get(n,'케어포'),'추천인코드':r.get('추천인코드',''),'피추천인수':0,'biz':[]}
        b=str(r.get('피추천인 사업자 번호','')).strip()
        if b and b not in ['-','nan']: ra[n]['피추천인수']+=1; ra[n]['biz'].append(b)
    b2u=members.set_index('사업자번호')['아이디'].to_dict(); bs=filtered.groupby('주문자 ID')['판매합계금액'].sum().to_dict()
    for n in ra: ra[n]['피추천인매출']=sum(bs.get(b2u.get(b,''),0) for b in ra[n]['biz'])
    rdf=pd.DataFrame(ra.values())[['추천인','유형','추천인코드','피추천인수','피추천인매출']]
    cols=st.columns(3)
    for col,(l,v,u) in zip(cols,[("총 추천인 수",fmt_num(len(rdf)),"회원"),("총 피추천인 수",fmt_num(rdf['피추천인수'].sum()),"회원"),("추천인당 평균 피추천인",f"{rdf['피추천인수'].mean():.1f}" if len(rdf)>0 else "0","회원")]): col.markdown(kpi_card(l,v,u),unsafe_allow_html=True)
    tc={'영업팀':'#3366CC','대리점':'#E8853D','케어포':'#27AE60'}
    cl,cr=st.columns(2)
    with cl:
        st.markdown("#### 추천인 유형별 피추천인 수")
        tr_df=rdf.groupby('유형')['피추천인수'].sum().reset_index()
        fig=px.bar(tr_df,x='유형',y='피추천인수',color='유형',color_discrete_map=tc)
        fig.update_traces(text=[fmt_num(v) for v in tr_df['피추천인수']],textposition='outside',textfont=dict(size=12),hovertemplate='%{x}: %{y:,}회원<extra></extra>')
        fig.update_layout(height=450,showlegend=False,margin=dict(l=60,r=30,t=30,b=40),xaxis=dict(title='',tickfont=dict(size=13)),yaxis=dict(title='피추천인 수 (회원)',tickfont=dict(size=11)))
        st.plotly_chart(fig, use_container_width=True)
    with cr:
        st.markdown("#### 추천인 유형별 피추천인 매출")
        ts_ref=rdf.groupby('유형')['피추천인매출'].sum().reset_index()
        fig = make_donut(ts_ref,'유형','피추천인매출',colors=[tc.get(t,'#999') for t in ts_ref['유형']])
        st.plotly_chart(fig, use_container_width=True)
    st.markdown("#### 추천인별 현황")
    rtf=st.selectbox("추천인 유형 필터",["전체","영업팀","대리점","케어포"],key="ref_type"); dr=rdf.copy()
    if rtf!="전체": dr=dr[dr['유형']==rtf]
    dr=dr.sort_values('피추천인매출',ascending=False)
    sr=st.text_input("🔍 추천인 검색",key="ref_search")
    if sr: dr=dr[dr.apply(lambda r:sr.lower() in str(r).lower(),axis=1)]
    st.dataframe(dr.style.format({'피추천인수':'{:,.0f}','피추천인매출':'{:,.0f}원'}),use_container_width=True,height=550)

# ============================================================
# Tab 6. 케어포 멤버십
# ============================================================
with tab6:
    cfg=['케어포-시설','케어포-공생','케어포-주야간','케어포-방문','케어포-일반','케어포-종사자','케어포-보호자']
    co=filtered[filtered['회원 등급'].isin(cfg)]; cmb=members[members['회원타입']=='케어포']; cf_filtered=filtered_members[filtered_members['회원타입']=='케어포']
    cgf=st.selectbox("케어포 등급",["전체"]+cfg,key="cf_grade")
    if cgf!="전체": co=co[co['회원 등급']==cgf]; cf_filtered=cf_filtered[cf_filtered['회원등급']==cgf]
    cbo=co.groupby('주문자 ID')['주문 ID'].nunique(); crp=(cbo>=2).sum(); crr=crp/len(cbo)*100 if len(cbo)>0 else 0
    cols=st.columns(4)
    for col,(l,v,u) in zip(cols,[("케어포 총 회원",fmt_num(len(cmb)),"처"),("케어포 신규가입",fmt_num(len(cf_filtered)),"처"),("케어포 주문회원",fmt_num(co['주문자 ID'].nunique()),"처"),("케어포 재구매율",fmt_pct(crr),"")]): col.markdown(kpi_card(l,v,u),unsafe_allow_html=True)
    grade_order = ['시설','공생','주야간','방문','일반','보호자','종사자']
    cga=co.groupby('회원 등급').agg(매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique')).reset_index(); cga['등급']=cga['회원 등급'].str.replace('케어포-','')
    cga['등급'] = pd.Categorical(cga['등급'], categories=grade_order, ordered=True); cga = cga.sort_values('등급')
    tvals_cf, ttexts_cf = krw_tickvals(cga['매출'])
    c1,c2 = st.columns(2)
    with c1:
        st.markdown("#### 케어포 등급별 매출")
        fig = px.bar(cga,x='등급',y='매출',color_discrete_sequence=['#3366CC'])
        fig.update_traces(text=[fmt_krw_short(v) for v in cga['매출']],textposition='outside',textfont=dict(size=11),hovertemplate='%{x}<br>매출: %{customdata}<extra></extra>',customdata=[f"{v:,.0f}원" for v in cga['매출']])
        fig.update_layout(height=420,margin=dict(l=60,r=20,t=30,b=40),showlegend=False,xaxis=dict(title='',tickfont=dict(size=12),categoryorder='array',categoryarray=grade_order),yaxis=dict(title='매출액',tickvals=tvals_cf,ticktext=ttexts_cf,tickfont=dict(size=11)))
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        st.markdown("#### 케어포 등급별 주문건수")
        fig = px.bar(cga,x='등급',y='주문건수',color_discrete_sequence=['#E8853D'])
        fig.update_traces(text=[fmt_num(v) for v in cga['주문건수']],textposition='outside',textfont=dict(size=11),hovertemplate='%{x}<br>주문: %{y:,}건<extra></extra>')
        fig.update_layout(height=420,margin=dict(l=60,r=20,t=30,b=40),showlegend=False,xaxis=dict(title='',tickfont=dict(size=12),categoryorder='array',categoryarray=grade_order),yaxis=dict(title='주문건수',tickfont=dict(size=11)))
        st.plotly_chart(fig, use_container_width=True)
    st.markdown("#### 케어포 월별 매출 추이")
    cf_monthly=co.groupby('주문월')['판매합계금액'].sum().reset_index(); cf_monthly['주문월_kr'] = ym_series_kr(cf_monthly['주문월'])
    tvals_cfm, ttexts_cfm = krw_tickvals(cf_monthly['판매합계금액'])
    fig=go.Figure()
    fig.add_trace(go.Bar(x=cf_monthly['주문월_kr'],y=cf_monthly['판매합계금액'],name='매출액',marker_color='#27AE60',opacity=0.8,text=[fmt_krw_short(v) for v in cf_monthly['판매합계금액']],textposition='outside',textfont=dict(size=10),hovertemplate='%{x}<br>매출: %{customdata}<extra></extra>',customdata=[f"{v:,.0f}원" for v in cf_monthly['판매합계금액']]))
    z = np.polyfit(range(len(cf_monthly)),cf_monthly['판매합계금액'].values,1); trend = np.polyval(z,range(len(cf_monthly)))
    fig.add_trace(go.Scatter(x=cf_monthly['주문월_kr'],y=trend,name='추세선',line=dict(color='#E74C3C',width=2,dash='dash'),mode='lines',hoverinfo='skip'))
    fig.update_layout(height=450,margin=dict(l=70,r=30,t=30,b=60),legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=11)),xaxis=dict(title='',tickfont=dict(size=12)),yaxis=dict(title='매출액',tickvals=tvals_cfm,ticktext=ttexts_cfm,tickfont=dict(size=11)))
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("#### 케어포 등급별 신규가입 추이")
    cj=cf_filtered.groupby(['가입월','회원등급']).size().reset_index(name='가입자수'); cj['가입월_kr'] = ym_series_kr(cj['가입월'])
    ccl={'케어포-시설':'#3366CC','케어포-공생':'#E8853D','케어포-주야간':'#27AE60','케어포-방문':'#9B59B6','케어포-일반':'#F39C12','케어포-종사자':'#E74C3C','케어포-보호자':'#1ABC9C'}
    fig=px.bar(cj,x='가입월_kr',y='가입자수',color='회원등급',color_discrete_map=ccl)
    for tr in fig.data: tr.hovertemplate = '%{x}<br>' + tr.name + ': %{y:,}처<extra></extra>'
    fig.update_layout(height=480,barmode='stack',margin=dict(l=60,r=20,t=30,b=100),legend=dict(orientation="h",yanchor="top",y=-0.12,x=0,font=dict(size=10)),xaxis=dict(title='',tickfont=dict(size=12)),yaxis=dict(title='가입자 수 (처)',tickfont=dict(size=11)))
    st.plotly_chart(fig, use_container_width=True)

# ============================================================
# Tab 7. 손익 분석 (BW)
# ============================================================
with tab7:
    if len(bw_data) == 0:
        st.warning("⚠️ BW 손익 데이터가 없습니다. 엑셀 파일에 'BW' 시트를 추가해주세요.")
    else:
        bw = bw_filtered.copy()

        # --- 손익 탭 전용 필터 ---
        bw_channels = sorted(bw['채널'].unique().tolist())
        sel_channels = st.multiselect("채널 필터", bw_channels, default=[], placeholder="전체", key="bw_channel")
        if sel_channels:
            bw = bw[bw['채널'].isin(sel_channels)]

        prod_large = sorted(bw['제품계층구조(대)'].dropna().unique().tolist())
        sel_prod_l = st.multiselect("제품계층구조(대) 필터", prod_large, default=[], placeholder="전체", key="bw_prod_l")
        if sel_prod_l:
            bw = bw[bw['제품계층구조(대)'].isin(sel_prod_l)]

        # --- KPI 카드 ---
        rev = bw['I.매출액(FI기준)'].sum()
        cogs = bw['II.매출원가'].sum()
        gp = bw['III.매출총이익'].sum()
        sga = bw['IV.판매비 및 관리비'].sum()
        oi = bw['V.영업이익I'].sum()
        oi_rate = (oi / rev * 100) if rev != 0 else 0

        cols = st.columns(6)
        kpis = [
            ("매출액", fmt_krw_short(rev), "원"),
            ("매출원가", fmt_krw_short(cogs), "원"),
            ("매출총이익", fmt_krw_short(gp), "원"),
            ("판관비", fmt_krw_short(sga), "원"),
            ("영업이익", fmt_krw_short(oi), "원"),
            ("영업이익률", fmt_pct(oi_rate), ""),
        ]
        for col, (l, v, u) in zip(cols, kpis):
            col.markdown(kpi_card(l, v, u), unsafe_allow_html=True)

        # --- 차트 1: 손익 워터폴 ---
        st.markdown("#### 손익 워터폴")
        wf_labels = ['매출액', '매출원가', '매출총이익', '판관비', '영업이익']
        wf_values = [rev, -cogs, gp, -sga, oi]
        wf_measure = ['absolute', 'relative', 'total', 'relative', 'total']
        wf_colors = ['#3366CC', '#E74C3C', '#27AE60', '#E74C3C', '#27AE60']
        fig = go.Figure(go.Waterfall(
            x=wf_labels, y=wf_values, measure=wf_measure,
            connector=dict(line=dict(color="#94a3b8", width=1)),
            increasing=dict(marker=dict(color='#27AE60')),
            decreasing=dict(marker=dict(color='#E74C3C')),
            totals=dict(marker=dict(color='#3366CC')),
            text=[fmt_krw_short(abs(v)) for v in wf_values],
            textposition='outside', textfont=dict(size=12),
            hovertemplate='%{x}<br>금액: %{customdata}<extra></extra>',
            customdata=[fmt_krw(abs(v)) for v in wf_values]
        ))
        wf_tvals, wf_ttexts = krw_tickvals(pd.Series([abs(v) for v in wf_values]))
        fig.update_layout(height=480, margin=dict(l=80, r=30, t=50, b=60),
            xaxis=dict(tickfont=dict(size=13)),
            yaxis=dict(title='금액', tickvals=wf_tvals, ticktext=wf_ttexts, tickfont=dict(size=11)),
            showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

        # --- 차트 2: 월별 손익 추이 ---
        st.markdown("#### 월별 손익 추이")
        bw_monthly = bw.groupby('연월').agg(
            매출액=('I.매출액(FI기준)', 'sum'),
            매출총이익=('III.매출총이익', 'sum'),
            영업이익=('V.영업이익I', 'sum')
        ).reset_index().sort_values('연월')
        bw_monthly['영업이익률'] = np.where(bw_monthly['매출액'] != 0, bw_monthly['영업이익'] / bw_monthly['매출액'] * 100, 0)
        bw_monthly['연월_kr'] = ym_series_kr(bw_monthly['연월'])

        fig = make_subplots(specs=[[{"secondary_y": True}]])
        for col_name, color, name in [('매출액', '#3366CC', '매출액'), ('매출총이익', '#27AE60', '매출총이익'), ('영업이익', '#E8853D', '영업이익')]:
            fig.add_trace(go.Bar(
                x=bw_monthly['연월_kr'], y=bw_monthly[col_name], name=name,
                marker_color=color, opacity=0.85,
                hovertemplate='%{x}<br>' + name + ': %{customdata}<extra></extra>',
                customdata=[fmt_krw(v) for v in bw_monthly[col_name]]
            ), secondary_y=False)
        fig.add_trace(go.Scatter(
            x=bw_monthly['연월_kr'], y=bw_monthly['영업이익률'], name='영업이익률',
            line=dict(color='#E74C3C', width=3), mode='lines+markers+text',
            marker=dict(size=8), text=[f"{v:.1f}%" for v in bw_monthly['영업이익률']],
            textposition='top center', textfont=dict(size=11, color='#E74C3C'),
            hovertemplate='%{x}<br>영업이익률: %{y:.1f}%<extra></extra>'
        ), secondary_y=True)
        fig.update_layout(height=500, barmode='group',
            margin=dict(l=80, r=60, t=50, b=70),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=11)),
            xaxis=dict(tickfont=dict(size=12)))
        monthly_tvals, monthly_ttexts = krw_tickvals(bw_monthly[['매출액','매출총이익','영업이익']].abs().max())
        fig.update_yaxes(title_text="금액", tickvals=monthly_tvals, ticktext=monthly_ttexts, tickfont=dict(size=11), secondary_y=False)
        fig.update_yaxes(title_text="영업이익률 (%)", tickfont=dict(size=11), ticksuffix='%', secondary_y=True)
        st.plotly_chart(fig, use_container_width=True)

        # --- 차트 3: 채널별 손익 비교 ---
        st.markdown("#### 채널별 손익 비교")
        ch_pnl = bw.groupby('채널').agg(
            매출액=('I.매출액(FI기준)', 'sum'),
            매출원가=('II.매출원가', 'sum'),
            매출총이익=('III.매출총이익', 'sum'),
            판관비=('IV.판매비 및 관리비', 'sum'),
            영업이익=('V.영업이익I', 'sum')
        ).reset_index()
        ch_pnl['매출총이익률'] = np.where(ch_pnl['매출액'] != 0, ch_pnl['매출총이익'] / ch_pnl['매출액'] * 100, 0)
        ch_pnl['영업이익률'] = np.where(ch_pnl['매출액'] != 0, ch_pnl['영업이익'] / ch_pnl['매출액'] * 100, 0)
        ch_pnl = ch_pnl.sort_values('매출액', ascending=True)

        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(
            x=ch_pnl['매출액'], y=ch_pnl['채널'], name='매출액', orientation='h',
            marker_color='#3366CC', opacity=0.8,
            text=[fmt_krw_short(v) for v in ch_pnl['매출액']], textposition='outside', textfont=dict(size=10),
            hovertemplate='%{y}<br>매출: %{customdata}<extra></extra>',
            customdata=[fmt_krw(v) for v in ch_pnl['매출액']]
        ), secondary_y=False)
        fig.add_trace(go.Scatter(
            x=ch_pnl['영업이익률'], y=ch_pnl['채널'], name='영업이익률', mode='markers+text',
            marker=dict(color='#E74C3C', size=12, symbol='diamond'),
            text=[f"{v:.1f}%" for v in ch_pnl['영업이익률']], textposition='middle right', textfont=dict(size=10, color='#E74C3C'),
            hovertemplate='%{y}<br>영업이익률: %{x:.1f}%<extra></extra>',
            xaxis='x2'
        ), secondary_y=False)
        ch_tvals, ch_ttexts = krw_tickvals(ch_pnl['매출액'])
        fig.update_layout(
            height=max(450, len(ch_pnl) * 38 + 140),
            margin=dict(l=130, r=80, t=30, b=40),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=11)),
            xaxis=dict(title='매출액', tickvals=ch_tvals, ticktext=ch_ttexts, tickfont=dict(size=11), side='bottom'),
            xaxis2=dict(title='영업이익률 (%)', tickfont=dict(size=11), side='top', overlaying='x', ticksuffix='%'),
            yaxis=dict(title='', tickfont=dict(size=11))
        )
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### 채널별 손익 상세")
        ch_display = ch_pnl.sort_values('매출액', ascending=False)[['채널','매출액','매출원가','매출총이익','매출총이익률','판관비','영업이익','영업이익률']].reset_index(drop=True)
        st.dataframe(ch_display.style.format({
            '매출액':'{:,.0f}원','매출원가':'{:,.0f}원','매출총이익':'{:,.0f}원',
            '매출총이익률':'{:.1f}%','판관비':'{:,.0f}원','영업이익':'{:,.0f}원','영업이익률':'{:.1f}%'
        }), use_container_width=True, height=450)

        # --- 차트 4: 판관비 구성 분석 ---
        st.markdown("#### 판관비 구성 분석")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("##### 판관비 항목별 비중")
            adv = bw['IV.6.광고선전비'].sum()
            freight = bw['IV.7.운반비'].sum()
            commission = bw['IV.8.판매수수료'].sum()
            promo = bw['IV.9.판촉비'].sum()
            etc_sga = sga - adv - freight - commission - promo
            sga_df = pd.DataFrame({
                '항목': ['광고선전비', '운반비', '판매수수료', '판촉비', '기타판관비'],
                '금액': [adv, freight, commission, promo, etc_sga]
            })
            sga_df = sga_df[sga_df['금액'] > 0].sort_values('금액', ascending=False)
            fig = make_donut(sga_df, '항목', '금액')
            fig.update_layout(height=480)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            st.markdown("##### 월별 판관비 구성 추이")
            sga_monthly = bw.groupby('연월').agg(
                광고선전비=('IV.6.광고선전비', 'sum'),
                운반비=('IV.7.운반비', 'sum'),
                판매수수료=('IV.8.판매수수료', 'sum'),
                판촉비=('IV.9.판촉비', 'sum')
            ).reset_index().sort_values('연월')
            sga_monthly['기타판관비'] = bw.groupby('연월')['IV.판매비 및 관리비'].sum().values - sga_monthly[['광고선전비','운반비','판매수수료','판촉비']].sum(axis=1).values
            sga_monthly['연월_kr'] = ym_series_kr(sga_monthly['연월'])
            sga_cols = ['광고선전비','운반비','판매수수료','판촉비','기타판관비']
            sga_colors = ['#3366CC','#E8853D','#27AE60','#9B59B6','#94a3b8']
            fig = go.Figure()
            for col_name, color in zip(sga_cols, sga_colors):
                fig.add_trace(go.Bar(
                    x=sga_monthly['연월_kr'], y=sga_monthly[col_name], name=col_name,
                    marker_color=color,
                    hovertemplate='%{x}<br>' + col_name + ': %{customdata}<extra></extra>',
                    customdata=[fmt_krw(v) for v in sga_monthly[col_name]]
                ))
            sga_tvals, sga_ttexts = krw_tickvals(sga_monthly[['광고선전비','운반비','판매수수료','판촉비','기타판관비']].sum(axis=1))
            fig.update_layout(height=480, barmode='stack',
                margin=dict(l=70, r=20, t=30, b=70),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=10)),
                xaxis=dict(tickfont=dict(size=12)),
                yaxis=dict(title='판관비', tickvals=sga_tvals, ticktext=sga_ttexts, tickfont=dict(size=11)))
            st.plotly_chart(fig, use_container_width=True)

        # --- 차트 5: 제품계층구조별 수익성 분석 (서브탭) ---
        st.markdown("#### 제품계층구조별 수익성 분석")
        prod_sub1, prod_sub2, prod_sub3 = st.tabs(["대분류", "중분류", "소분류"])

        def render_product_pnl(df, group_col, tab_key):
            """제품계층구조별 손익 차트 + 테이블 렌더링"""
            pnl = df.groupby(group_col).agg(
                매출액=('I.매출액(FI기준)', 'sum'),
                매출총이익=('III.매출총이익', 'sum'),
                영업이익=('V.영업이익I', 'sum')
            ).reset_index()
            pnl['매출총이익률'] = np.where(pnl['매출액'] != 0, pnl['매출총이익'] / pnl['매출액'] * 100, 0)
            pnl['영업이익률'] = np.where(pnl['매출액'] != 0, pnl['영업이익'] / pnl['매출액'] * 100, 0)
            pnl = pnl.sort_values('매출액', ascending=True)

            fig = make_subplots(specs=[[{"secondary_y": True}]])
            fig.add_trace(go.Bar(
                x=pnl['매출액'], y=pnl[group_col], name='매출액', orientation='h',
                marker_color='#3366CC', opacity=0.8,
                text=[fmt_krw_short(v) for v in pnl['매출액']], textposition='outside', textfont=dict(size=10),
                hovertemplate='%{y}<br>매출: %{customdata}<extra></extra>',
                customdata=[fmt_krw(v) for v in pnl['매출액']]
            ), secondary_y=False)
            fig.add_trace(go.Scatter(
                x=pnl['영업이익률'], y=pnl[group_col], name='영업이익률', mode='markers+text',
                marker=dict(color='#E74C3C', size=10, symbol='diamond'),
                text=[f"{v:.1f}%" for v in pnl['영업이익률']], textposition='middle right', textfont=dict(size=10, color='#E74C3C'),
                hovertemplate='%{y}<br>영업이익률: %{x:.1f}%<extra></extra>',
                xaxis='x2'
            ), secondary_y=False)
            pnl_tvals, pnl_ttexts = krw_tickvals(pnl['매출액'])
            fig.update_layout(
                height=max(420, len(pnl) * 30 + 140),
                margin=dict(l=180, r=80, t=30, b=40),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=11)),
                xaxis=dict(title='매출액', tickvals=pnl_tvals, ticktext=pnl_ttexts, tickfont=dict(size=11)),
                xaxis2=dict(title='영업이익률 (%)', tickfont=dict(size=11), side='top', overlaying='x', ticksuffix='%'),
                yaxis=dict(title='', tickfont=dict(size=10))
            )
            st.plotly_chart(fig, use_container_width=True)

            # 테이블
            tbl = pnl.sort_values('매출액', ascending=False).reset_index(drop=True)
            search_key = f"prod_search_{tab_key}"
            ps = st.text_input(f"🔍 {group_col} 검색", key=search_key)
            if ps:
                tbl = tbl[tbl[group_col].str.contains(ps, case=False, na=False)]
            st.dataframe(tbl.style.format({
                '매출액':'{:,.0f}원','매출총이익':'{:,.0f}원','매출총이익률':'{:.1f}%',
                '영업이익':'{:,.0f}원','영업이익률':'{:.1f}%'
            }), use_container_width=True, height=400)

        with prod_sub1:
            render_product_pnl(bw, '제품계층구조(대)', 'large')

        with prod_sub2:
            # 대분류 필터 연동
            sel_large = st.selectbox("대분류 선택", ["전체"] + sorted(bw['제품계층구조(대)'].unique().tolist()), key="bw_mid_filter")
            bw_mid = bw if sel_large == "전체" else bw[bw['제품계층구조(대)'] == sel_large]
            render_product_pnl(bw_mid, '제품계층구조(중)', 'medium')

        with prod_sub3:
            # 대분류 + 중분류 필터 연동
            c1, c2 = st.columns(2)
            with c1:
                sel_large2 = st.selectbox("대분류 선택", ["전체"] + sorted(bw['제품계층구조(대)'].unique().tolist()), key="bw_small_filter_l")
            bw_small = bw if sel_large2 == "전체" else bw[bw['제품계층구조(대)'] == sel_large2]
            with c2:
                mid_opts = sorted(bw_small['제품계층구조(중)'].unique().tolist())
                sel_mid = st.selectbox("중분류 선택", ["전체"] + mid_opts, key="bw_small_filter_m")
            if sel_mid != "전체":
                bw_small = bw_small[bw_small['제품계층구조(중)'] == sel_mid]
            render_product_pnl(bw_small, '제품계층구조(소)', 'small')

        # --- 차트 6: 자재별 손익 테이블 ---
        st.markdown("#### 자재별 손익 현황")
        mat_pnl = bw.groupby(['자재', '자재명']).agg(
            매출액=('I.매출액(FI기준)', 'sum'),
            매출원가=('II.매출원가', 'sum'),
            매출총이익=('III.매출총이익', 'sum'),
            판관비=('IV.판매비 및 관리비', 'sum'),
            영업이익=('V.영업이익I', 'sum'),
            판매수량=('판매수량', 'sum')
        ).reset_index()
        mat_pnl['매출총이익률'] = np.where(mat_pnl['매출액'] != 0, mat_pnl['매출총이익'] / mat_pnl['매출액'] * 100, 0)
        mat_pnl['영업이익률'] = np.where(mat_pnl['매출액'] != 0, mat_pnl['영업이익'] / mat_pnl['매출액'] * 100, 0)
        mat_pnl = mat_pnl.sort_values('매출액', ascending=False).reset_index(drop=True)

        ms = st.text_input("🔍 자재명/코드 검색", key="bw_mat_search")
        if ms:
            mat_pnl = mat_pnl[mat_pnl.apply(lambda r: ms.lower() in str(r['자재명']).lower() or ms in str(r['자재']), axis=1)]

        def highlight_negative(val):
            if isinstance(val, (int, float)) and val < 0:
                return 'color: #E74C3C; font-weight: 600'
            return ''

        st.dataframe(
            mat_pnl.style.format({
                '매출액':'{:,.0f}원','매출원가':'{:,.0f}원','매출총이익':'{:,.0f}원',
                '매출총이익률':'{:.1f}%','판관비':'{:,.0f}원','영업이익':'{:,.0f}원',
                '영업이익률':'{:.1f}%','판매수량':'{:,.0f}'
            }).map(highlight_negative, subset=['영업이익','영업이익률']),
            use_container_width=True, height=550
        )

# ============================================================
# 푸터
# ============================================================
st.markdown("---")
st.markdown(f"<p style='text-align:center;color:#94a3b8;font-size:0.85rem;'>© 대상웰라이프 B2B몰 대시보드 · 데이터 기준: {pd.Timestamp.now().strftime('%Y년 %m월 %d일')}</p>", unsafe_allow_html=True)
