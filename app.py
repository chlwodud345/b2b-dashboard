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
# 호버 전역 설정
HOVER_FONT = dict(font=dict(size=16, family='Noto Sans KR'))

# ============================================================
# CSS
# ============================================================
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
.js-plotly-plot .plotly .hoverlayer .hovertext text { font-size: 22px !important; }
.js-plotly-plot .plotly .hoverlayer .hovertext path { stroke-width: 2px !important; }
div[data-testid="stVerticalBlock"] > div { padding-top: 4px; padding-bottom: 4px; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 포맷팅 함수
# ============================================================
def fmt_krw(n):
    if pd.isna(n) or n == 0: return "0원"
    if abs(n) >= 1e8: return f"{n/1e8:.1f}억원"
    if abs(n) >= 1e4: return f"{n/1e4:,.0f}만원"
    return f"{n:,.0f}원"

def fmt_krw_short(n):
    if pd.isna(n) or n == 0: return "0"
    if abs(n) >= 1e8: return f"{n/1e8:.1f}억"
    if abs(n) >= 1e4: return f"{n/1e4:,.0f}만"
    return f"{n:,.0f}"

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
    """'2025-02-15' → '2025년 2월 15일'"""
    if not date_str or pd.isna(date_str): return ''
    parts = str(date_str).split('-')
    if len(parts) == 3: return f"{parts[0]}년 {int(parts[1])}월 {int(parts[2])}일"
    return str(date_str)

def krw_tickvals(series, n=5):
    """y축 tick을 한국어 원화로 표시하기 위한 값/텍스트 생성"""
    mn, mx = series.min(), series.max()
    if mx == 0: return [0], ['0']
    vals = np.linspace(0, mx * 1.05, n).tolist()
    texts = [fmt_krw_short(v) for v in vals]
    return vals, texts

def make_donut(df, name_col, value_col, title, colors=None):
    """도넛 차트 - 가운데 합계, 범례를 별도 영역에 배치"""
    total = df[value_col].sum()
    fig = go.Figure()
    fig.add_trace(go.Pie(
        labels=df[name_col], values=df[value_col],
        hole=0.55, marker=dict(colors=(colors or COLORS)[:len(df)]),
        textinfo='label+percent', textposition='inside',
        insidetextorientation='horizontal',
        textfont=dict(size=11),
        hovertemplate='%{label}<br>매출: %{customdata}<br>비중: %{percent}<extra></extra>',
        customdata=[fmt_krw(v) for v in df[value_col]],
    ))
    fig.add_annotation(
        text=f"<b>합계</b><br>{fmt_krw(total)}",
        x=0.5, y=0.5, font=dict(size=15, color='#1e293b'),
        showarrow=False, xref='paper', yref='paper'
    )
    fig.update_layout(
        height=520, title=dict(text=title, x=0.01, font=dict(size=16)),
        margin=dict(l=20, r=20, t=90, b=140),
        legend=dict(orientation="h", yanchor="top", y=-0.02, xanchor="center", x=0.5,
                    font=dict(size=11), traceorder='normal'),
        showlegend=True
    )
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

# ============================================================
# 데이터 로드
# ============================================================
GDRIVE_FILE_ID = '1Op9Y2FFb_aLQJKAcLyKj9HJQbK6YYnmf'

def download_from_gdrive(file_id):
    session = requests.Session()
    url = f'https://drive.google.com/uc?export=download&id={file_id}'
    response = session.get(url, stream=True)
    confirm_token = None
    for key, value in response.cookies.items():
        if key.startswith('download_warning'): confirm_token = value; break
    if confirm_token:
        response = session.get(f'https://drive.google.com/uc?export=download&confirm={confirm_token}&id={file_id}', stream=True)
    content = response.content
    if content[:4] != b'PK\x03\x04':
        response = session.get(f'https://drive.google.com/uc?export=download&confirm=t&id={file_id}', stream=True)
        content = response.content
    return io.BytesIO(content)

@st.cache_data(ttl=3600, show_spinner="📥 구글 드라이브에서 데이터를 불러오는 중...")
def load_from_gdrive():
    fb = download_from_gdrive(GDRIVE_FILE_ID)
    o = pd.read_excel(fb, sheet_name='주문내역', header=1, engine='openpyxl'); fb.seek(0)
    m = pd.read_excel(fb, sheet_name='회원정보', header=1, engine='openpyxl'); fb.seek(0)
    r = pd.read_excel(fb, sheet_name='추천인', header=1, engine='openpyxl')
    return process_data(o, m, r)

try:
    orders, members, referrals_df = load_from_gdrive()
    st.sidebar.success(f"✅ 데이터 로드 완료\n- 주문: {len(orders):,}건\n- 회원: {len(members):,}건\n- 추천인: {len(referrals_df):,}건")
except Exception as e:
    st.error(f"❌ 데이터 로드 실패: {str(e)}\n\n구글 드라이브 파일 공유 설정을 확인해주세요.")
    st.stop()

if st.sidebar.button("🔄 데이터 새로고침"):
    st.cache_data.clear()
    st.rerun()
st.sidebar.markdown("---")
member_lookup = members.set_index('아이디')[['상호명','회원타입','회원등급']].to_dict('index')

# ============================================================
# 사이드바 필터
# ============================================================
st.sidebar.markdown("## 🔍 필터")
years = sorted(orders['주문일'].dt.year.dropna().unique().astype(int))
selected_year = st.sidebar.selectbox("연도", ["전체"] + [str(y) for y in years], index=0)
if selected_year != "전체":
    mo_opts = sorted(orders[orders['주문일'].dt.year == int(selected_year)]['주문일'].dt.month.dropna().unique().astype(int))
    selected_month = st.sidebar.selectbox("월", ["전체"] + [f"{m}월" for m in mo_opts], index=0)
else:
    selected_month = "전체"
member_types = ["전체"] + sorted(orders['주문자 구분'].dropna().unique().tolist())
selected_type = st.sidebar.selectbox("회원구분", member_types, index=0)
member_grades = ["전체"] + sorted(orders['회원 등급'].dropna().unique().tolist())
selected_grade = st.sidebar.selectbox("회원등급", member_grades, index=0)

filtered = orders.copy()
if selected_year != "전체": filtered = filtered[filtered['주문일'].dt.year == int(selected_year)]
if selected_month != "전체": filtered = filtered[filtered['주문일'].dt.month == int(selected_month.replace('월',''))]
if selected_type != "전체": filtered = filtered[filtered['주문자 구분'] == selected_type]
if selected_grade != "전체": filtered = filtered[filtered['회원 등급'] == selected_grade]

filtered_members = members.copy()
if selected_year != "전체": filtered_members = filtered_members[filtered_members['가입일'].dt.year == int(selected_year)]
if selected_month != "전체": filtered_members = filtered_members[filtered_members['가입일'].dt.month == int(selected_month.replace('월',''))]
if selected_type != "전체": filtered_members = filtered_members[filtered_members['회원타입'] == selected_type]
if selected_grade != "전체": filtered_members = filtered_members[filtered_members['회원등급'] == selected_grade]

# ============================================================
# 헤더 & 탭
# ============================================================
st.markdown('<div class="main-header"><h1>📊 대상웰라이프 B2B몰 대시보드</h1><p>Sales & Operations Analytics</p></div>', unsafe_allow_html=True)
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["📋 종합 현황","💰 매출 분석","📦 상품 분석","👥 회원 분석","🔗 추천인 분석","💚 케어포 멤버십"])

# ============================================================
# Tab 1. 종합 현황
# ============================================================
with tab1:
    ts = filtered['판매합계금액'].sum()
    to_ = filtered['주문 ID'].nunique()
    tb = filtered['주문자 ID'].nunique()
    tm = len(members); nm = len(filtered_members)
    ao = ts/to_ if to_>0 else 0
    
    cols = st.columns(6)
    for col,(l,v,u) in zip(cols,[
        ("총 매출액",fmt_krw_short(ts),"원"),("총 주문건수",fmt_num(to_),"건"),
        ("총 회원수",fmt_num(tm),"처"),("주문회원수",fmt_num(tb),"처"),
        ("신규 가입회원",fmt_num(nm),"처"),("객단가",fmt_krw_short(ao),"원"),
    ]):
        col.markdown(kpi_card(l,v,u), unsafe_allow_html=True)
    
    # 월별 매출·주문건수 (전체 너비)
    monthly = filtered.groupby('주문월').agg(매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique')).reset_index()
    monthly['주문월_kr'] = ym_series_kr(monthly['주문월'])
    fig = make_subplots(specs=[[{"secondary_y":True}]])
    fig.add_trace(go.Bar(x=monthly['주문월_kr'],y=monthly['매출'],name='매출액',marker_color='#3366CC',opacity=0.8,
                         text=[fmt_krw_short(v) for v in monthly['매출']],textposition='outside',textfont=dict(size=11),
                         hovertemplate='%{x}<br>매출: %{customdata}<extra></extra>',
                         customdata=[fmt_krw(v) for v in monthly['매출']]),secondary_y=False)
    fig.add_trace(go.Scatter(x=monthly['주문월_kr'],y=monthly['주문건수'],name='주문건수',
                             line=dict(color='#E8853D',width=3),mode='lines+markers',marker=dict(size=8),
                             hovertemplate='%{x}<br>주문: %{y:,}건<extra></extra>'),secondary_y=True)
    tvals, ttexts = krw_tickvals(monthly['매출'])
    fig.update_layout(height=480,margin=dict(l=80,r=60,t=100,b=70),
                      title=dict(text='월별 매출 · 주문건수 추이',x=0.01,font=dict(size=17)),
                      legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=12)),
                      xaxis=dict(tickfont=dict(size=12)))
    fig.update_yaxes(title_text="매출액",tickvals=tvals,ticktext=ttexts,tickfont=dict(size=11),secondary_y=False)
    fig.update_yaxes(title_text="주문건수",tickfont=dict(size=11),secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)
    
    # 회원구분별 매출 + 지역별 매출 (한 줄)
    cl,cr = st.columns(2)
    with cl:
        ts_df = filtered.groupby('주문자 구분')['판매합계금액'].sum().reset_index()
        ts_df.columns = ['구분','매출']
        ts_df = ts_df.sort_values('매출',ascending=False)
        fig = make_donut(ts_df, '구분', '매출', '회원구분별 매출 비중')
        fig.update_layout(height=520)
        st.plotly_chart(fig, use_container_width=True)
    with cr:
        rg = filtered.groupby('지역')['판매합계금액'].sum().sort_values().reset_index()
        rg.columns = ['지역','매출']
        fig = px.bar(rg,x='매출',y='지역',orientation='h',color_discrete_sequence=COLORS)
        fig.update_traces(text=[fmt_krw_short(v) for v in rg['매출']],textposition='outside',textfont=dict(size=11),
                          hovertemplate='%{y}: %{customdata}<extra></extra>',customdata=[fmt_krw(v) for v in rg['매출']])
        fig.update_layout(height=520,title=dict(text='지역별 매출',x=0.01,font=dict(size=17)),
                          margin=dict(l=70,r=100,t=90,b=40),showlegend=False,
                          xaxis=dict(title='',tickfont=dict(size=11)),yaxis=dict(title='',tickfont=dict(size=12)))
        st.plotly_chart(fig, use_container_width=True)
    
    # 일별 매출 (전체 너비)
    daily = filtered.groupby('주문일자')['판매합계금액'].sum().reset_index()
    daily.columns = ['날짜','매출']
    daily['날짜_kr'] = daily['날짜'].apply(to_date_kr)
    fig = px.area(daily,x='날짜',y='매출',color_discrete_sequence=['#3366CC'])
    fig.update_traces(
        customdata=list(zip(daily['날짜'].apply(to_date_kr), [fmt_krw(v) for v in daily['매출']])),
        hovertemplate='%{customdata[0]}<br>매출: %{customdata[1]}<extra></extra>'
    )
    tvals2, ttexts2 = krw_tickvals(daily['매출'])
    fig.update_layout(height=400,title=dict(text='일별 매출 추이',x=0.01,font=dict(size=17)),
                      margin=dict(l=80,r=30,t=100,b=60),showlegend=False,
                      xaxis=dict(title='날짜',tickfont=dict(size=11),title_font=dict(size=13),tickformat='%Y년 %m월'),
                      yaxis=dict(title='매출액',tickvals=tvals2,ticktext=ttexts2,tickfont=dict(size=11),title_font=dict(size=13)))
    st.plotly_chart(fig, use_container_width=True)

# ============================================================
# Tab 2. 매출 분석
# ============================================================
with tab2:
    # 회원구분별 × 월별 매출
    tm_df = filtered.groupby(['주문월','주문자 구분'])['판매합계금액'].sum().reset_index()
    tm_df['주문월_kr'] = ym_series_kr(tm_df['주문월'])
    fig = px.bar(tm_df,x='주문월_kr',y='판매합계금액',color='주문자 구분',color_discrete_sequence=COLORS)
    for tr in fig.data:
        tr.customdata = [fmt_krw(v) for v in tr.y]
        tr.hovertemplate = '%{x}<br>' + tr.name + ': %{customdata}<extra></extra>'
    fig.update_layout(height=480,barmode='stack',
                      title=dict(text='회원구분별 × 월별 매출 추이',x=0.01,font=dict(size=17)),
                      margin=dict(l=70,r=30,t=130,b=70),
                      legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=11)),
                      xaxis=dict(title='',tickfont=dict(size=12)),yaxis=dict(title='매출액',tickfont=dict(size=11)))
    st.plotly_chart(fig, use_container_width=True)
    
    # 회원등급별 매출
    gs = filtered.groupby('회원 등급').agg(매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique'),주문회원수=('주문자 ID','nunique')).reset_index().sort_values('매출')
    fig = px.bar(gs,x='매출',y='회원 등급',orientation='h',color_discrete_sequence=COLORS)
    fig.update_traces(text=[fmt_krw_short(v) for v in gs['매출']],textposition='outside',textfont=dict(size=11),
                      hovertemplate='%{y}<br>매출: %{customdata[0]}<br>주문: %{customdata[1]:,}건<br>회원: %{customdata[2]:,}처<extra></extra>',
                      customdata=list(zip([fmt_krw(v) for v in gs['매출']], gs['주문건수'], gs['주문회원수'])))
    fig.update_layout(height=max(420,len(gs)*35+140),
                      title=dict(text='회원등급별 매출',x=0.01,font=dict(size=17)),
                      margin=dict(l=140,r=100,t=100,b=40),showlegend=False,
                      xaxis=dict(title='매출액',tickfont=dict(size=11)),yaxis=dict(title='',tickfont=dict(size=12)))
    st.plotly_chart(fig, use_container_width=True)
    
    # 히트맵
    st.markdown("#### 요일 · 시간대별 주문 매출 히트맵")
    dow_order = ['월','화','수','목','금','토','일']
    hm = filtered.groupby(['요일','주문시간'])['판매합계금액'].sum().reset_index()
    hmp = hm.pivot_table(index='요일',columns='주문시간',values='판매합계금액',fill_value=0).reindex(dow_order)
    fig = go.Figure(data=go.Heatmap(z=hmp.values,x=[f'{h}시' for h in hmp.columns],y=hmp.index,
                                      colorscale=[[0,'#F0F2F5'],[0.5,'#6B9BD2'],[1,'#1B2A4A']],
                                      text=[[fmt_krw_short(v) for v in row] for row in hmp.values],
                                      texttemplate='%{text}',textfont=dict(size=12),
                                      hovertemplate='%{y} %{x}<br>매출: %{customdata}<extra></extra>',
                                      customdata=[[f"{v:,.0f}원" for v in row] for row in hmp.values]))
    fig.update_layout(height=320,margin=dict(l=50,r=20,t=30,b=40),
                      xaxis=dict(tickfont=dict(size=11)),yaxis=dict(tickfont=dict(size=12),autorange='reversed'))
    st.plotly_chart(fig, use_container_width=True)
    
    # 기관별 매출 테이블
    st.markdown("#### 기관별 매출 현황")
    ba = filtered.groupby(['주문자 ID','주문자명','주문자 구분','회원 등급']).agg(
        매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique'),최근주문일=('주문일자','max')).reset_index()
    ba['객단가'] = (ba['매출']/ba['주문건수']).round(0)
    ba['상호명'] = ba['주문자 ID'].map(lambda x: member_lookup.get(x, {}).get('상호명', ''))
    ba = ba[['주문자 ID','주문자명','상호명','주문자 구분','회원 등급','주문건수','매출','객단가','최근주문일']]
    ba = ba.sort_values('매출',ascending=False)
    search = st.text_input("🔍 검색 (아이디, 주문자명, 상호명)",key="sales_search")
    if search: ba = ba[ba.apply(lambda r:search.lower() in str(r).lower(),axis=1)]
    st.dataframe(ba.style.format({'매출':'{:,.0f}원','주문건수':'{:,.0f}건','객단가':'{:,.0f}원'}),
                 use_container_width=True,height=550)

# ============================================================
# Tab 3. 상품 분석
# ============================================================
with tab3:
    pa = filtered.groupby(['상품명','상품 코드']).agg(매출=('판매합계금액','sum'),수량=('주문 수량','sum'),주문건수=('주문 ID','nunique')).reset_index().sort_values('매출',ascending=False)
    
    top20 = pa.head(20).copy()
    ttl = pa['매출'].sum()
    top20['누적비중'] = (top20['매출'].cumsum()/ttl*100).round(1)
    
    fig = make_subplots(specs=[[{"secondary_y":True}]])
    fig.add_trace(go.Bar(x=[f"{i+1}. {n[:16]}" for i,n in enumerate(top20['상품명'])],y=top20['매출'],name='매출액',marker_color='#3366CC',opacity=0.8,
                         hovertemplate='%{customdata[0]}<br>매출: %{customdata[1]}<extra></extra>',
                         customdata=list(zip(top20['상품명'],[fmt_krw(v) for v in top20['매출']]))),secondary_y=False)
    fig.add_trace(go.Scatter(x=[f"{i+1}. {n[:16]}" for i,n in enumerate(top20['상품명'])],y=top20['누적비중'],name='누적 비중',
                             line=dict(color='#E8853D',width=3),mode='lines+markers',marker=dict(size=7),
                             hovertemplate='%{customdata}<br>누적 비중: %{y:.1f}%<extra></extra>',
                             customdata=top20['상품명'].tolist()),secondary_y=True)
    tvals3, ttexts3 = krw_tickvals(top20['매출'])
    fig.update_layout(height=540,title=dict(text='상품별 매출 TOP 20 (파레토 차트)',x=0.01,font=dict(size=17)),
                      margin=dict(l=80,r=60,t=100,b=150),
                      legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=12)),
                      xaxis=dict(tickangle=45,tickfont=dict(size=9)))
    fig.update_yaxes(title_text="매출액",tickvals=tvals3,ticktext=ttexts3,tickfont=dict(size=11),secondary_y=False)
    fig.update_yaxes(title_text="누적 비중 (%)",range=[0,100],ticksuffix='%',tickfont=dict(size=11),secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("#### 전체 상품 매출 현황")
    sp = st.text_input("🔍 상품명/코드 검색",key="product_search")
    dp = pa.copy()
    if sp: dp = dp[dp.apply(lambda r:sp.lower() in str(r).lower(),axis=1)]
    st.dataframe(dp.style.format({'매출':'{:,.0f}원','수량':'{:,.0f}','주문건수':'{:,.0f}건'}),
                 use_container_width=True,height=450)
    
    st.markdown("#### 회원구분별 × 상품 매출 크로스 (TOP 20)")
    t20n = pa.head(20)['상품명'].tolist()
    cp = filtered[filtered['상품명'].isin(t20n)].pivot_table(index='상품명',columns='주문자 구분',values='판매합계금액',aggfunc='sum',fill_value=0)
    cp['합계'] = cp.sum(axis=1); cp = cp.sort_values('합계',ascending=False)
    st.dataframe(cp.style.format('{:,.0f}원'),use_container_width=True,height=550)
    
    st.markdown("#### 월별 상품 매출 추이")
    t5 = pa.head(5)['상품명'].tolist()
    sel = st.multiselect("상품 선택",pa['상품명'].tolist(),default=t5,key="product_trend")
    if sel:
        td = filtered[filtered['상품명'].isin(sel)].groupby(['주문월','상품명'])['판매합계금액'].sum().reset_index()
        td['주문월_kr'] = ym_series_kr(td['주문월'])
        fig = px.line(td,x='주문월_kr',y='판매합계금액',color='상품명',markers=True,color_discrete_sequence=COLORS)
        for tr in fig.data:
            tr.customdata = [fmt_krw(v) for v in tr.y]
            tr.hovertemplate = '%{x}<br>' + (tr.name[:20]+'...' if len(tr.name)>20 else tr.name) + '<br>매출: %{customdata}<extra></extra>'
            if len(tr.name) > 22: tr.name = tr.name[:22]+'...'
        fig.update_layout(height=480,margin=dict(l=70,r=30,t=50,b=120),
                          legend=dict(orientation="h",yanchor="top",y=-0.15,x=0,font=dict(size=10)),
                          xaxis=dict(title='',tickfont=dict(size=12)),yaxis=dict(title='매출액',tickfont=dict(size=11)))
        st.plotly_chart(fig, use_container_width=True)

# ============================================================
# Tab 4. 회원 분석
# ============================================================
with tab4:
    mo_df = orders.groupby('주문자 ID').agg(첫주문일=('주문일','min'),주문건수=('주문 ID','nunique'),주문월수=('주문월','nunique')).reset_index()
    conv = members[members['아이디'].isin(orders['주문자 ID'].unique())]
    conv_r = len(conv)/len(members)*100 if len(members)>0 else 0
    rep = mo_df[mo_df['주문건수']>=2]; rep_r = len(rep)/len(mo_df)*100 if len(mo_df)>0 else 0
    r3m = orders[orders['주문일']>=orders['주문일'].max()-pd.DateOffset(months=3)]; act = r3m['주문자 ID'].nunique()
    
    cols = st.columns(5)
    for col,(l,v,u) in zip(cols,[
        ("총 회원수",fmt_num(len(members)),"처"),("신규 가입회원",fmt_num(len(filtered_members)),"처"),
        ("구매전환율",fmt_pct(conv_r),""),("재구매율",fmt_pct(rep_r),""),("활성회원(3개월)",fmt_num(act),"처"),
    ]):
        col.markdown(kpi_card(l,v,u),unsafe_allow_html=True)
    
    # ---- 회원 상세 검색 ----
    st.markdown("#### 🔍 회원 상세 검색")
    mem_search = st.text_input("상호명, 아이디, 담당자명으로 검색", key="member_detail_search", placeholder="예: 뉴케어, cs3, 홍길동")
    
    if mem_search:
        # 회원정보 검색
        mask = members.apply(lambda r: mem_search.lower() in str(r.get('상호명','')).lower()
                             or mem_search.lower() in str(r.get('아이디','')).lower()
                             or mem_search.lower() in str(r.get('담당자 이름','')).lower(), axis=1)
        search_members = members[mask].copy()
        
        if len(search_members) == 0:
            st.warning("검색 결과가 없습니다.")
        else:
            # 주문 집계 조인
            order_agg = orders.groupby('주문자 ID').agg(
                총매출=('판매합계금액','sum'),
                주문건수=('주문 ID','nunique'),
                첫주문일=('주문일자','min'),
                최근주문일=('주문일자','max')
            ).reset_index()
            order_agg['객단가'] = (order_agg['총매출']/order_agg['주문건수']).round(0)
            
            # 주요 구매상품 (매출 TOP 3)
            top_products = orders.groupby(['주문자 ID','상품명'])['판매합계금액'].sum().reset_index()
            top_products = top_products.sort_values(['주문자 ID','판매합계금액'],ascending=[True,False])
            top_products = top_products.groupby('주문자 ID').head(3)
            top_products = top_products.groupby('주문자 ID')['상품명'].apply(lambda x: ' / '.join(x)).reset_index()
            top_products.columns = ['주문자 ID','주요 구매상품']
            
            # 추천인 정보 조인 (사업자번호 기준)
            ref_info = referrals_df.groupby('피추천인 사업자 번호').first()[['추천인','회원그룹']].reset_index()
            ref_info.columns = ['사업자번호','추천인명','추천인유형']
            
            # 결합
            result = search_members[['아이디','상호명','사업자번호','회원타입','회원등급','가입일','담당자 이름','휴대폰','주소']].copy()
            result['가입일'] = result['가입일'].dt.strftime('%Y-%m-%d')
            result = result.merge(order_agg, left_on='아이디', right_on='주문자 ID', how='left').drop(columns=['주문자 ID'],errors='ignore')
            result = result.merge(top_products, left_on='아이디', right_on='주문자 ID', how='left').drop(columns=['주문자 ID'],errors='ignore')
            result = result.merge(ref_info, on='사업자번호', how='left')
            
            # 정리
            result['총매출'] = result['총매출'].fillna(0)
            result['주문건수'] = result['주문건수'].fillna(0).astype(int)
            result['객단가'] = result['객단가'].fillna(0)
            result['주요 구매상품'] = result['주요 구매상품'].fillna('-')
            result['추천인명'] = result['추천인명'].fillna('-')
            result['추천인유형'] = result['추천인유형'].fillna('-')
            
            display_cols = ['아이디','상호명','담당자 이름','회원타입','회원등급','가입일',
                           '총매출','주문건수','객단가','첫주문일','최근주문일','주요 구매상품',
                           '추천인명','추천인유형','휴대폰','주소']
            result = result[[c for c in display_cols if c in result.columns]]
            result = result.sort_values('총매출',ascending=False)
            
            st.markdown(f"**검색 결과: {len(result)}건**")
            st.dataframe(
                result.style.format({'총매출':'{:,.0f}원','주문건수':'{:,.0f}건','객단가':'{:,.0f}원'}),
                use_container_width=True, height=400
            )
    
    # 월별 신규가입 (전체 너비)
    jm = filtered_members.groupby(['가입월','회원타입']).size().reset_index(name='가입자수')
    jm['가입월_kr'] = ym_series_kr(jm['가입월'])
    fig = px.bar(jm,x='가입월_kr',y='가입자수',color='회원타입',color_discrete_sequence=COLORS)
    for tr in fig.data:
        tr.hovertemplate = '%{x}<br>' + tr.name + ': %{y:,}처<extra></extra>'
    fig.update_layout(height=480,barmode='stack',
                      title=dict(text='월별 신규가입자 추이 (회원타입별)',x=0.01,font=dict(size=17)),
                      margin=dict(l=60,r=30,t=130,b=70),
                      legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=11)),
                      xaxis=dict(title='',tickfont=dict(size=12)),yaxis=dict(title='가입자 수 (처)',tickfont=dict(size=11)))
    st.plotly_chart(fig, use_container_width=True)
    
    # 회원등급별 분포 (전체 너비)
    gd = filtered_members['회원등급'].value_counts().reset_index(); gd.columns = ['등급','수']
    fig = px.bar(gd,x='수',y='등급',orientation='h',color_discrete_sequence=COLORS)
    fig.update_traces(text=[fmt_num(v) for v in gd['수']],textposition='outside',textfont=dict(size=11),
                      hovertemplate='%{y}: %{x:,}처<extra></extra>')
    fig.update_layout(height=max(420,len(gd)*35+140),title=dict(text='회원등급별 가입자 분포',x=0.01,font=dict(size=17)),
                      margin=dict(l=140,r=80,t=100,b=40),showlegend=False,
                      xaxis=dict(title='가입자 수 (처)',tickfont=dict(size=11)),yaxis=dict(title='',tickfont=dict(size=12)))
    st.plotly_chart(fig, use_container_width=True)
    
    cl,cr = st.columns(2)
    with cl:
        mg = filtered_members.merge(mo_df,left_on='아이디',right_on='주문자 ID',how='inner')
        mg['소요일'] = (mg['첫주문일']-mg['가입일']).dt.days; mg = mg[mg['소요일']>=0]
        bins=[0,1,8,15,31,61,91,9999]; lb=['당일','1~7일','8~14일','15~30일','31~60일','61~90일','90일+']
        mg['구간'] = pd.cut(mg['소요일'],bins=bins,labels=lb,right=False)
        dh = mg['구간'].value_counts().reindex(lb).fillna(0).reset_index(); dh.columns=['구간','회원수']
        fig = px.bar(dh,x='구간',y='회원수',color_discrete_sequence=['#3366CC'])
        fig.update_traces(text=[fmt_num(v) for v in dh['회원수']],textposition='outside',textfont=dict(size=11),
                          hovertemplate='%{x}: %{y:,}처<extra></extra>')
        fig.update_layout(height=450,title=dict(text='가입 후 첫 주문까지 소요일',x=0.01,font=dict(size=16)),
                          margin=dict(l=60,r=30,t=100,b=60),showlegend=False,
                          xaxis=dict(title='',tickfont=dict(size=12)),yaxis=dict(title='회원 수 (처)',tickfont=dict(size=11)))
        st.plotly_chart(fig, use_container_width=True)
    with cr:
        mo_df2 = mo_df.copy()
        bo=[1,2,4,6,11,21,9999]; lo=['1회','2~3회','4~5회','6~10회','11~20회','20회+']
        mo_df2['구간'] = pd.cut(mo_df2['주문건수'],bins=bo,labels=lo,right=False)
        od = mo_df2['구간'].value_counts().reindex(lo).fillna(0).reset_index(); od.columns=['구간','회원수']
        fig = px.bar(od,x='구간',y='회원수',color_discrete_sequence=COLORS)
        fig.update_traces(text=[fmt_num(v) for v in od['회원수']],textposition='outside',textfont=dict(size=11),
                          hovertemplate='%{x}: %{y:,}처<extra></extra>')
        fig.update_layout(height=450,title=dict(text='주문횟수 구간별 회원 분포',x=0.01,font=dict(size=16)),
                          margin=dict(l=60,r=30,t=100,b=60),showlegend=False,
                          xaxis=dict(title='',tickfont=dict(size=12)),yaxis=dict(title='회원 수 (처)',tickfont=dict(size=11)))
        st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("#### 코호트 리텐션 히트맵 (가입월 × 경과월별 재구매율)")
    cm = members[members['가입일'].notna()].copy()
    cm['코호트'] = cm['가입일'].dt.to_period('M').astype(str)
    om = orders.groupby('주문자 ID')['주문월'].apply(set).to_dict()
    cohorts = sorted(cm['코호트'].unique()); mx=12; rd=[]
    for c in cohorts:
        cu = cm[cm['코호트']==c]['아이디'].tolist(); sz=len(cu)
        if sz==0: continue
        row={'코호트':c,'크기':sz}; cp=pd.Period(c,freq='M')
        for o in range(mx):
            t=str(cp+o); a=sum(1 for u in cu if t in om.get(u,set()))
            row[f'{o}개월']=round(a/sz*100,1)
        rd.append(row)
    if rd:
        rdf=pd.DataFrame(rd); mc=[f'{o}개월' for o in range(mx)]; zv=rdf[mc].values
        fig=go.Figure(data=go.Heatmap(z=zv,x=mc,
            y=[f"{to_ym_kr(r['코호트'])} ({r['크기']}처)" for _,r in rdf.iterrows()],
            colorscale=[[0,'#F0F2F5'],[0.3,'#A8D5A2'],[1,'#27AE60']],
            text=[[f'{v:.1f}%' if v>0 else '-' for v in row] for row in zv],
            texttemplate='%{text}',textfont=dict(size=10),
            hovertemplate='%{y}<br>%{x}: %{z:.1f}%<extra></extra>'))
        fig.update_layout(height=max(400,len(rd)*32+140),margin=dict(l=190,r=20,t=30,b=50),
                          yaxis=dict(tickfont=dict(size=10),autorange="reversed"),
                          xaxis=dict(tickfont=dict(size=11)))
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
    b2u=members.set_index('사업자번호')['아이디'].to_dict()
    bs=filtered.groupby('주문자 ID')['판매합계금액'].sum().to_dict()
    for n in ra: ra[n]['피추천인매출']=sum(bs.get(b2u.get(b,''),0) for b in ra[n]['biz'])
    rdf=pd.DataFrame(ra.values())[['추천인','유형','추천인코드','피추천인수','피추천인매출']]
    
    cols=st.columns(3)
    for col,(l,v,u) in zip(cols,[
        ("총 추천인 수",fmt_num(len(rdf)),"회원"),("총 피추천인 수",fmt_num(rdf['피추천인수'].sum()),"회원"),
        ("추천인당 평균 피추천인",f"{rdf['피추천인수'].mean():.1f}" if len(rdf)>0 else "0","회원"),
    ]): col.markdown(kpi_card(l,v,u),unsafe_allow_html=True)
    
    tc={'영업팀':'#3366CC','대리점':'#E8853D','케어포':'#27AE60'}
    cl,cr=st.columns(2)
    with cl:
        tr_df=rdf.groupby('유형')['피추천인수'].sum().reset_index()
        fig=px.bar(tr_df,x='유형',y='피추천인수',color='유형',color_discrete_map=tc)
        fig.update_traces(text=[fmt_num(v) for v in tr_df['피추천인수']],textposition='outside',textfont=dict(size=12),
                          hovertemplate='%{x}: %{y:,}회원<extra></extra>')
        fig.update_layout(height=450,showlegend=False,title=dict(text='추천인 유형별 피추천인 수',x=0.01,font=dict(size=17)),
                          margin=dict(l=60,r=30,t=100,b=40),
                          xaxis=dict(title='',tickfont=dict(size=13)),yaxis=dict(title='피추천인 수 (회원)',tickfont=dict(size=11)))
        st.plotly_chart(fig, use_container_width=True)
    with cr:
        ts_ref=rdf.groupby('유형')['피추천인매출'].sum().reset_index()
        fig = make_donut(ts_ref,'유형','피추천인매출','추천인 유형별 피추천인 매출',
                         colors=[tc.get(t,'#999') for t in ts_ref['유형']])
        st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("#### 추천인별 현황")
    rtf=st.selectbox("추천인 유형 필터",["전체","영업팀","대리점","케어포"],key="ref_type")
    dr=rdf.copy()
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
    co=filtered[filtered['회원 등급'].isin(cfg)]; cmb=members[members['회원타입']=='케어포']
    cf_filtered=filtered_members[filtered_members['회원타입']=='케어포']
    cgf=st.selectbox("케어포 등급",["전체"]+cfg,key="cf_grade")
    if cgf!="전체": co=co[co['회원 등급']==cgf]; cf_filtered=cf_filtered[cf_filtered['회원등급']==cgf]
    cbo=co.groupby('주문자 ID')['주문 ID'].nunique()
    crp=(cbo>=2).sum(); crr=crp/len(cbo)*100 if len(cbo)>0 else 0
    
    cols=st.columns(4)
    for col,(l,v,u) in zip(cols,[
        ("케어포 총 회원",fmt_num(len(cmb)),"처"),("케어포 신규가입",fmt_num(len(cf_filtered)),"처"),
        ("케어포 주문회원",fmt_num(co['주문자 ID'].nunique()),"처"),("케어포 재구매율",fmt_pct(crr),""),
    ]): col.markdown(kpi_card(l,v,u),unsafe_allow_html=True)
    
    # 등급별 매출·주문
    cga=co.groupby('회원 등급').agg(매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique')).reset_index()
    cga['등급']=cga['회원 등급'].str.replace('케어포-','')
    fig=make_subplots(specs=[[{"secondary_y":True}]])
    fig.add_trace(go.Bar(x=cga['등급'],y=cga['매출'],name='매출액',marker_color='#3366CC',opacity=0.8,
                         text=[fmt_krw_short(v) for v in cga['매출']],textposition='outside',textfont=dict(size=10),
                         hovertemplate='%{x}<br>매출: %{customdata}<extra></extra>',customdata=[fmt_krw(v) for v in cga['매출']]),secondary_y=False)
    fig.add_trace(go.Scatter(x=cga['등급'],y=cga['주문건수'],name='주문건수',
                             line=dict(color='#E8853D',width=3),mode='lines+markers',marker=dict(size=8),
                             hovertemplate='%{x}<br>주문: %{y:,}건<extra></extra>'),secondary_y=True)
    fig.update_layout(height=480,title=dict(text='케어포 등급별 매출 · 주문',x=0.01,font=dict(size=17)),
                      margin=dict(l=70,r=60,t=100,b=40),
                      legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=11)),
                      xaxis=dict(title='',tickfont=dict(size=13)))
    fig.update_yaxes(title_text="매출액",tickfont=dict(size=11),secondary_y=False)
    fig.update_yaxes(title_text="주문건수",tickfont=dict(size=11),secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)
    
    # 전용 상품 매출
    cpd=co[co['상품명'].str.contains(r'\[케어포',na=False)]
    cpm=cpd.groupby('주문월')['판매합계금액'].sum().reset_index()
    cpm['주문월_kr'] = ym_series_kr(cpm['주문월'])
    fig=px.area(cpm,x='주문월_kr',y='판매합계금액',color_discrete_sequence=['#27AE60'])
    fig.update_traces(hovertemplate='%{x}<br>매출: %{customdata}<extra></extra>',customdata=[fmt_krw(v) for v in cpm['판매합계금액']])
    fig.update_layout(height=450,title=dict(text='케어포 전용 상품 매출 추이',x=0.01,font=dict(size=17)),
                      margin=dict(l=70,r=30,t=100,b=60),showlegend=False,
                      xaxis=dict(title='',tickfont=dict(size=12)),yaxis=dict(title='매출액',tickfont=dict(size=11)))
    st.plotly_chart(fig, use_container_width=True)
    
    # 신규가입 추이
    cj=cf_filtered.groupby(['가입월','회원등급']).size().reset_index(name='가입자수')
    cj['가입월_kr'] = ym_series_kr(cj['가입월'])
    ccl={'케어포-시설':'#3366CC','케어포-공생':'#E8853D','케어포-주야간':'#27AE60',
         '케어포-방문':'#9B59B6','케어포-일반':'#F39C12','케어포-종사자':'#E74C3C','케어포-보호자':'#1ABC9C'}
    fig=px.bar(cj,x='가입월_kr',y='가입자수',color='회원등급',color_discrete_map=ccl)
    for tr in fig.data: tr.hovertemplate = '%{x}<br>' + tr.name + ': %{y:,}처<extra></extra>'
    fig.update_layout(height=480,barmode='stack',
                      title=dict(text='케어포 등급별 신규가입 추이',x=0.01,font=dict(size=17)),
                      margin=dict(l=60,r=20,t=100,b=100),
                      legend=dict(orientation="h",yanchor="top",y=-0.12,x=0,font=dict(size=10)),
                      xaxis=dict(title='',tickfont=dict(size=12)),yaxis=dict(title='가입자 수 (처)',tickfont=dict(size=11)))
    st.plotly_chart(fig, use_container_width=True)

st.markdown("---")
st.markdown(f"<p style='text-align:center;color:#94a3b8;font-size:0.85rem;'>© 대상웰라이프 B2B몰 대시보드 · 데이터 기준: {pd.Timestamp.now().strftime('%Y년 %m월 %d일')}</p>",unsafe_allow_html=True)
