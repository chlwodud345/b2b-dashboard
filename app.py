import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import requests

# ============================================================
# 페이지 설정
# ============================================================
st.set_page_config(page_title="대상웰라이프 B2B몰 대시보드", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

COLORS = ['#3366CC','#E8853D','#27AE60','#9B59B6','#E74C3C','#1ABC9C','#F39C12','#2980B9','#8E44AD','#D35400']

# ============================================================
# CSS
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700&display=swap');
html, body, [class*="st-"] { font-family: 'Noto Sans KR', sans-serif; }
.main-header { background: linear-gradient(135deg, #1B2A4A 0%, #2D4A7A 100%); color: white; padding: 20px 30px; border-radius: 12px; margin-bottom: 24px; }
.main-header h1 { margin: 0; font-size: 1.5rem; font-weight: 700; }
.main-header p { margin: 4px 0 0; opacity: 0.7; font-size: 0.85rem; }
.kpi-card { background: white; border-radius: 12px; padding: 20px; border: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0,0,0,0.04); text-align: center; margin-bottom: 16px; }
.kpi-label { font-size: 0.8rem; color: #64748b; font-weight: 500; margin-bottom: 4px; }
.kpi-value { font-size: 1.6rem; font-weight: 700; color: #1e293b; }
.kpi-unit { font-size: 0.85rem; color: #94a3b8; margin-left: 2px; }
.stTabs [data-baseweb="tab-list"] { gap: 4px; }
.stTabs [data-baseweb="tab"] { padding: 10px 20px; font-weight: 500; }
[data-testid="stSidebar"] { background: #f8fafc; }
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
    """'2025-02' → '2025년 2월'"""
    if not ym_str or pd.isna(ym_str): return ''
    parts = str(ym_str).split('-')
    if len(parts) >= 2:
        return f"{parts[0]}년 {int(parts[1])}월"
    return str(ym_str)

def ym_series_kr(series):
    """시리즈 전체를 한국어 연월로 변환"""
    return series.apply(to_ym_kr)

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
# 데이터 로드 (구글 드라이브)
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

# 회원 룩업 테이블 (상호명 조인용) - GAS의 memberLookup과 동일
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

# 필터 적용 (주문)
filtered = orders.copy()
if selected_year != "전체": filtered = filtered[filtered['주문일'].dt.year == int(selected_year)]
if selected_month != "전체": filtered = filtered[filtered['주문일'].dt.month == int(selected_month.replace('월',''))]
if selected_type != "전체": filtered = filtered[filtered['주문자 구분'] == selected_type]
if selected_grade != "전체": filtered = filtered[filtered['회원 등급'] == selected_grade]

# 필터 적용 (회원) - GAS의 filterMembers와 동일
filtered_members = members.copy()
if selected_year != "전체":
    filtered_members = filtered_members[filtered_members['가입일'].dt.year == int(selected_year)]
if selected_month != "전체":
    m_val = int(selected_month.replace('월',''))
    filtered_members = filtered_members[filtered_members['가입일'].dt.month == m_val]
if selected_type != "전체":
    filtered_members = filtered_members[filtered_members['회원타입'] == selected_type]
if selected_grade != "전체":
    filtered_members = filtered_members[filtered_members['회원등급'] == selected_grade]

# ============================================================
# 헤더 & 탭
# ============================================================
st.markdown('<div class="main-header"><h1>📊 대상웰라이프 B2B몰 대시보드</h1><p>Sales & Operations Analytics</p></div>', unsafe_allow_html=True)
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["📋 종합 현황","💰 매출 분석","📦 상품 분석","👥 회원 분석","🔗 추천인 분석","💚 케어포 멤버십"])

# ============================================================
# Tab 1. 종합 현황 (GAS overview() 동일)
# ============================================================
with tab1:
    ts = filtered['판매합계금액'].sum()
    to_ = filtered['주문 ID'].nunique()
    tb = filtered['주문자 ID'].nunique()
    tm = len(members)
    nm = len(filtered_members)  # GAS: newMembers = filteredMembers.length
    ao = ts/to_ if to_>0 else 0
    
    cols = st.columns(6)
    for col,(l,v,u) in zip(cols,[
        ("총 매출액",fmt_krw_short(ts),"원"),("총 주문건수",fmt_num(to_),"건"),
        ("총 회원수",fmt_num(tm),"처"),("주문회원수",fmt_num(tb),"처"),
        ("신규 가입회원",fmt_num(nm),"처"),("객단가",fmt_krw_short(ao),"원"),
    ]):
        col.markdown(kpi_card(l,v,u), unsafe_allow_html=True)
    
    # 월별 매출 · 주문건수
    monthly = filtered.groupby('주문월').agg(매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique'),주문회원수=('주문자 ID','nunique')).reset_index()
    monthly['주문월_kr'] = ym_series_kr(monthly['주문월'])
    
    fig = make_subplots(specs=[[{"secondary_y":True}]])
    fig.add_trace(go.Bar(x=monthly['주문월_kr'],y=monthly['매출'],name='매출액',marker_color='#3366CC',opacity=0.8,
                         hovertemplate='%{x}<br>매출: %{customdata}<extra></extra>',
                         customdata=[fmt_krw(v) for v in monthly['매출']]),secondary_y=False)
    fig.add_trace(go.Scatter(x=monthly['주문월_kr'],y=monthly['주문건수'],name='주문건수',
                             line=dict(color='#E8853D',width=2.5),mode='lines+markers',
                             hovertemplate='%{x}<br>주문: %{y:,}건<extra></extra>'),secondary_y=True)
    fig.update_layout(height=420,margin=dict(l=60,r=50,t=70,b=60),
                      title=dict(text='월별 매출 · 주문건수 추이',x=0.01,font=dict(size=15)),
                      legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=11)))
    fig.update_yaxes(title_text="매출액",secondary_y=False)
    fig.update_yaxes(title_text="주문건수",secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)
    
    cl,cr = st.columns(2)
    with cl:
        # 회원구분별 매출 비중 - 전체 항목 표시 (기타 합침 없음)
        ts_df = filtered.groupby('주문자 구분')['판매합계금액'].sum().reset_index()
        ts_df.columns = ['구분','매출']
        ts_df = ts_df.sort_values('매출',ascending=False)
        fig = px.pie(ts_df, values='매출', names='구분', hole=0.5, color_discrete_sequence=COLORS)
        fig.update_traces(textinfo='label+percent', textfont_size=10,
                          hovertemplate='%{label}<br>매출: %{customdata}<br>비중: %{percent}<extra></extra>',
                          customdata=[fmt_krw(v) for v in ts_df['매출']])
        fig.update_layout(height=450,title=dict(text='회원구분별 매출 비중',x=0.01,font=dict(size=15)),
                          margin=dict(l=20,r=150,t=70,b=20),
                          legend=dict(orientation="v",yanchor="middle",y=0.5,xanchor="left",x=1.02,font=dict(size=10)))
        st.plotly_chart(fig, use_container_width=True)
    with cr:
        rg = filtered.groupby('지역')['판매합계금액'].sum().sort_values().reset_index()
        rg.columns = ['지역','매출']
        fig = px.bar(rg,x='매출',y='지역',orientation='h',color_discrete_sequence=COLORS)
        fig.update_traces(hovertemplate='%{y}: %{customdata}<extra></extra>',
                          customdata=[fmt_krw(v) for v in rg['매출']])
        fig.update_layout(height=450,title=dict(text='지역별 매출',x=0.01,font=dict(size=15)),
                          margin=dict(l=60,r=30,t=70,b=30),showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    daily = filtered.groupby('주문일자')['판매합계금액'].sum().reset_index()
    daily.columns = ['날짜','매출']
    fig = px.area(daily,x='날짜',y='매출',color_discrete_sequence=['#3366CC'])
    fig.update_traces(hovertemplate='%{x}<br>매출: %{customdata}<extra></extra>',
                      customdata=[fmt_krw(v) for v in daily['매출']])
    fig.update_layout(height=350,title=dict(text='일별 매출 추이',x=0.01,font=dict(size=15)),
                      margin=dict(l=60,r=30,t=70,b=50),showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

# ============================================================
# Tab 2. 매출 분석 (GAS sales() 동일)
# ============================================================
with tab2:
    # 회원구분별 × 월별 매출 (GAS: salesByTypeMonth)
    tm_df = filtered.groupby(['주문월','주문자 구분'])['판매합계금액'].sum().reset_index()
    tm_df['주문월_kr'] = ym_series_kr(tm_df['주문월'])
    fig = px.bar(tm_df,x='주문월_kr',y='판매합계금액',color='주문자 구분',color_discrete_sequence=COLORS)
    for tr in fig.data:
        tr.customdata = [fmt_krw(v) for v in tr.y]
        tr.hovertemplate = '%{x}<br>' + tr.name + ': %{customdata}<extra></extra>'
    fig.update_layout(height=420,barmode='stack',
                      title=dict(text='회원구분별 × 월별 매출 추이',x=0.01,font=dict(size=15)),
                      margin=dict(l=60,r=30,t=70,b=60),
                      legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=10)))
    st.plotly_chart(fig, use_container_width=True)
    
    # 회원등급별 매출 (GAS: salesByGrade)
    gs = filtered.groupby('회원 등급').agg(매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique'),
                                           주문회원수=('주문자 ID','nunique')).reset_index()
    gs = gs.sort_values('매출')
    fig = px.bar(gs,x='매출',y='회원 등급',orientation='h',color_discrete_sequence=COLORS)
    fig.update_traces(hovertemplate='%{y}<br>매출: %{customdata[0]}<br>주문: %{customdata[1]:,}건<br>회원: %{customdata[2]:,}처<extra></extra>',
                      customdata=list(zip([fmt_krw(v) for v in gs['매출']], gs['주문건수'], gs['주문회원수'])))
    fig.update_layout(height=max(350,len(gs)*28+100),
                      title=dict(text='회원등급별 매출',x=0.01,font=dict(size=15)),
                      margin=dict(l=130,r=30,t=70,b=30),showlegend=False)
    st.plotly_chart(fig, use_container_width=True)
    
    # 요일·시간대 히트맵 (GAS: dowHourHeatmap)
    st.markdown("##### 요일 · 시간대별 주문 매출 히트맵")
    dow_order = ['월','화','수','목','금','토','일']
    hm = filtered.groupby(['요일','주문시간'])['판매합계금액'].sum().reset_index()
    hmp = hm.pivot_table(index='요일',columns='주문시간',values='판매합계금액',fill_value=0).reindex(dow_order)
    hm_text = [[fmt_krw_short(v) for v in row] for row in hmp.values]
    fig = go.Figure(data=go.Heatmap(z=hmp.values,x=[f'{h}시' for h in hmp.columns],y=hmp.index,
                                      colorscale=[[0,'#F0F2F5'],[0.5,'#6B9BD2'],[1,'#1B2A4A']],
                                      text=hm_text,texttemplate='%{text}',textfont=dict(size=8),
                                      hovertemplate='%{y} %{x}<br>매출: %{customdata}<extra></extra>',
                                      customdata=[[fmt_krw(v) for v in row] for row in hmp.values]))
    fig.update_layout(height=300,margin=dict(l=50,r=20,t=20,b=40))
    st.plotly_chart(fig, use_container_width=True)
    
    # 기관별 매출 테이블 (GAS: buyerTable - 상호명 포함)
    st.markdown("##### 기관별 매출 현황")
    ba = filtered.groupby(['주문자 ID','주문자명','주문자 구분','회원 등급']).agg(
        매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique'),최근주문일=('주문일자','max')).reset_index()
    ba['객단가'] = (ba['매출']/ba['주문건수']).round(0)
    # 상호명 조인 (GAS의 memberLookup과 동일)
    ba['상호명'] = ba['주문자 ID'].map(lambda x: member_lookup.get(x, {}).get('상호명', ''))
    ba = ba[['주문자 ID','주문자명','상호명','주문자 구분','회원 등급','주문건수','매출','객단가','최근주문일']]
    ba = ba.sort_values('매출',ascending=False)
    search = st.text_input("🔍 검색 (아이디, 주문자명, 상호명)",key="sales_search")
    if search: ba = ba[ba.apply(lambda r:search.lower() in str(r).lower(),axis=1)]
    st.dataframe(ba.style.format({'매출':'{:,.0f}원','주문건수':'{:,.0f}건','객단가':'{:,.0f}원'}),
                 use_container_width=True,height=500)

# ============================================================
# Tab 3. 상품 분석 (GAS products() 동일)
# ============================================================
with tab3:
    pa = filtered.groupby(['상품명','상품 코드']).agg(매출=('판매합계금액','sum'),수량=('주문 수량','sum'),주문건수=('주문 ID','nunique')).reset_index().sort_values('매출',ascending=False)
    
    # 파레토 (GAS: productTable 상위 20개)
    top20 = pa.head(20).copy()
    ttl = pa['매출'].sum()
    top20['누적비중'] = (top20['매출'].cumsum()/ttl*100).round(1)
    
    fig = make_subplots(specs=[[{"secondary_y":True}]])
    fig.add_trace(go.Bar(x=top20['상품명'].str[:18],y=top20['매출'],name='매출액',marker_color='#3366CC',opacity=0.8,
                         hovertemplate='%{customdata[0]}<br>매출: %{customdata[1]}<extra></extra>',
                         customdata=list(zip(top20['상품명'],[fmt_krw(v) for v in top20['매출']]))),secondary_y=False)
    fig.add_trace(go.Scatter(x=top20['상품명'].str[:18],y=top20['누적비중'],name='누적 비중',
                             line=dict(color='#E8853D',width=2.5),mode='lines+markers',
                             hovertemplate='누적 비중: %{y:.1f}%<extra></extra>'),secondary_y=True)
    fig.update_layout(height=480,title=dict(text='상품별 매출 TOP 20 (파레토)',x=0.01,font=dict(size=15)),
                      margin=dict(l=60,r=50,t=70,b=130),
                      legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=10)),
                      xaxis=dict(tickangle=45,tickfont=dict(size=8)))
    fig.update_yaxes(title_text="매출액",secondary_y=False)
    fig.update_yaxes(title_text="누적 비중 (%)",range=[0,105],secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)
    
    # 전체 상품 테이블
    st.markdown("##### 전체 상품 매출 현황")
    sp = st.text_input("🔍 상품명/코드 검색",key="product_search")
    dp = pa.copy()
    if sp: dp = dp[dp.apply(lambda r:sp.lower() in str(r).lower(),axis=1)]
    st.dataframe(dp.style.format({'매출':'{:,.0f}원','수량':'{:,.0f}','주문건수':'{:,.0f}건'}),
                 use_container_width=True,height=400)
    
    # 크로스 분석 (GAS: productCross)
    st.markdown("##### 회원구분별 × 상품 매출 크로스 (TOP 20)")
    t20n = pa.head(20)['상품명'].tolist()
    cp = filtered[filtered['상품명'].isin(t20n)].pivot_table(index='상품명',columns='주문자 구분',values='판매합계금액',aggfunc='sum',fill_value=0)
    cp['합계'] = cp.sum(axis=1)
    cp = cp.sort_values('합계',ascending=False)
    st.dataframe(cp.style.format('{:,.0f}원'),use_container_width=True,height=500)
    
    # 월별 상품 추이 (GAS: productMonthly)
    st.markdown("##### 월별 상품 매출 추이")
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
        fig.update_layout(height=400,margin=dict(l=60,r=30,t=30,b=80),
                          legend=dict(orientation="h",yanchor="top",y=-0.18,x=0,font=dict(size=9)))
        st.plotly_chart(fig, use_container_width=True)

# ============================================================
# Tab 4. 회원 분석 (GAS membersTab() 동일)
# ============================================================
with tab4:
    # 회원별 주문 집계 (GAS: memberOrderMap)
    mo_df = orders.groupby('주문자 ID').agg(첫주문일=('주문일','min'),주문건수=('주문 ID','nunique'),주문월수=('주문월','nunique')).reset_index()
    
    # KPI (GAS: membersTab kpi)
    conv = members[members['아이디'].isin(orders['주문자 ID'].unique())]
    conv_r = len(conv)/len(members)*100 if len(members)>0 else 0
    rep = mo_df[mo_df['주문건수']>=2]
    rep_r = len(rep)/len(mo_df)*100 if len(mo_df)>0 else 0
    r3m = orders[orders['주문일']>=orders['주문일'].max()-pd.DateOffset(months=3)]
    act = r3m['주문자 ID'].nunique()
    
    cols = st.columns(5)
    for col,(l,v,u) in zip(cols,[
        ("총 회원수",fmt_num(len(members)),"처"),
        ("신규 가입회원",fmt_num(len(filtered_members)),"처"),
        ("구매전환율",fmt_pct(conv_r),""),
        ("재구매율",fmt_pct(rep_r),""),
        ("활성회원(3개월)",fmt_num(act),"처"),
    ]):
        col.markdown(kpi_card(l,v,u),unsafe_allow_html=True)
    
    # 월별 신규가입 (GAS: newMembersByMonth - 필터된 회원 기준)
    cl,cr = st.columns(2)
    with cl:
        jm = filtered_members.groupby(['가입월','회원타입']).size().reset_index(name='가입자수')
        jm['가입월_kr'] = ym_series_kr(jm['가입월'])
        fig = px.bar(jm,x='가입월_kr',y='가입자수',color='회원타입',color_discrete_sequence=COLORS)
        for tr in fig.data:
            tr.hovertemplate = '%{x}<br>' + tr.name + ': %{y:,}처<extra></extra>'
        fig.update_layout(height=420,barmode='stack',
                          title=dict(text='월별 신규가입자 추이 (회원타입별)',x=0.01,font=dict(size=15)),
                          margin=dict(l=50,r=20,t=70,b=60),
                          legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=10)))
        st.plotly_chart(fig, use_container_width=True)
    with cr:
        # 회원등급별 분포 (GAS: gradeDistribution - 필터된 회원 기준)
        gd = filtered_members['회원등급'].value_counts().reset_index()
        gd.columns = ['등급','수']
        fig = px.bar(gd,x='수',y='등급',orientation='h',color_discrete_sequence=COLORS)
        fig.update_traces(hovertemplate='%{y}: %{x:,}처<extra></extra>')
        fig.update_layout(height=420,title=dict(text='회원등급별 가입자 분포',x=0.01,font=dict(size=15)),
                          margin=dict(l=130,r=20,t=70,b=30),showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    # 첫 주문 소요일 (GAS: daysToFirstOrder)
    cl,cr = st.columns(2)
    with cl:
        mg = filtered_members.merge(mo_df,left_on='아이디',right_on='주문자 ID',how='inner')
        mg['소요일'] = (mg['첫주문일']-mg['가입일']).dt.days
        mg = mg[mg['소요일']>=0]
        bins=[0,1,8,15,31,61,91,9999]; lb=['당일','1~7일','8~14일','15~30일','31~60일','61~90일','90일+']
        mg['구간'] = pd.cut(mg['소요일'],bins=bins,labels=lb,right=False)
        dh = mg['구간'].value_counts().reindex(lb).fillna(0).reset_index()
        dh.columns=['구간','회원수']
        fig = px.bar(dh,x='구간',y='회원수',color_discrete_sequence=['#3366CC'])
        fig.update_traces(hovertemplate='%{x}: %{y:,}처<extra></extra>')
        fig.update_layout(height=400,title=dict(text='가입 후 첫 주문까지 소요일',x=0.01,font=dict(size=15)),
                          margin=dict(l=50,r=20,t=70,b=50),showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    with cr:
        # 주문횟수 분포 (GAS: orderCountDistribution)
        bo=[1,2,4,6,11,21,9999]; lo=['1회','2~3회','4~5회','6~10회','11~20회','20회+']
        mo_df2 = mo_df.copy()
        mo_df2['구간'] = pd.cut(mo_df2['주문건수'],bins=bo,labels=lo,right=False)
        od = mo_df2['구간'].value_counts().reindex(lo).fillna(0).reset_index()
        od.columns=['구간','회원수']
        fig = px.bar(od,x='구간',y='회원수',color_discrete_sequence=COLORS)
        fig.update_traces(hovertemplate='%{x}: %{y:,}처<extra></extra>')
        fig.update_layout(height=400,title=dict(text='주문횟수 구간별 회원 분포',x=0.01,font=dict(size=15)),
                          margin=dict(l=50,r=20,t=70,b=50),showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    # 코호트 리텐션 (GAS: cohortRetention - 전체 회원 기준)
    st.markdown("##### 코호트 리텐션 히트맵 (가입월 × 경과월별 재구매율)")
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
            texttemplate='%{text}',textfont=dict(size=9),
            hovertemplate='%{y}<br>%{x}: %{z:.1f}%<extra></extra>'))
        fig.update_layout(height=max(350,len(rd)*30+120),margin=dict(l=180,r=20,t=20,b=40),
                          yaxis=dict(tickfont=dict(size=9),autorange="reversed"))
        st.plotly_chart(fig, use_container_width=True)

# ============================================================
# Tab 5. 추천인 분석 (GAS referralsTab() 동일)
# ============================================================
with tab5:
    # 추천인 분류 (GAS: getRecommenderClassification)
    clm={}
    for _,r in referrals_df.iterrows():
        n=str(r.get('추천인','')).strip(); g=str(r.get('회원그룹',''))
        if not n or n in ['-','nan']: continue
        if g=='영업팀': clm[n]='케어포' if n=='케어포' else '영업팀'
        elif g=='대리점 회원': clm[n]='대리점'
    
    # 추천인별 피추천인 집계 (GAS: recMap)
    ra={}
    for _,r in referrals_df.iterrows():
        n=str(r.get('추천인','')).strip()
        if not n or n in ['-','nan']: continue
        if n not in ra:
            ra[n]={'추천인':n,'유형':clm.get(n,'케어포'),'추천인코드':r.get('추천인코드',''),'피추천인수':0,'biz':[]}
        b=str(r.get('피추천인 사업자 번호','')).strip()
        if b and b not in ['-','nan']:
            ra[n]['피추천인수']+=1; ra[n]['biz'].append(b)
    
    # 피추천인 매출 (GAS: refereeSales)
    b2u=members.set_index('사업자번호')['아이디'].to_dict()
    bs=filtered.groupby('주문자 ID')['판매합계금액'].sum().to_dict()
    for n in ra:
        ra[n]['피추천인매출']=sum(bs.get(b2u.get(b,''),0) for b in ra[n]['biz'])
    
    rdf=pd.DataFrame(ra.values())[['추천인','유형','추천인코드','피추천인수','피추천인매출']]
    
    # KPI
    cols=st.columns(3)
    for col,(l,v,u) in zip(cols,[
        ("총 추천인 수",fmt_num(len(rdf)),"회원"),
        ("총 피추천인 수",fmt_num(rdf['피추천인수'].sum()),"회원"),
        ("추천인당 평균 피추천인",f"{rdf['피추천인수'].mean():.1f}" if len(rdf)>0 else "0","회원"),
    ]):
        col.markdown(kpi_card(l,v,u),unsafe_allow_html=True)
    
    tc={'영업팀':'#3366CC','대리점':'#E8853D','케어포':'#27AE60'}
    cl,cr=st.columns(2)
    with cl:
        # 유형별 피추천인 수 (GAS: refereeByType)
        tr_df=rdf.groupby('유형')['피추천인수'].sum().reset_index()
        fig=px.bar(tr_df,x='유형',y='피추천인수',color='유형',color_discrete_map=tc)
        fig.update_traces(hovertemplate='%{x}: %{y:,}회원<extra></extra>')
        fig.update_layout(height=400,showlegend=False,
                          title=dict(text='추천인 유형별 피추천인 수',x=0.01,font=dict(size=15)),
                          margin=dict(l=50,r=20,t=70,b=30))
        st.plotly_chart(fig, use_container_width=True)
    with cr:
        # 유형별 매출 비중 (GAS: salesByType)
        ts_ref=rdf.groupby('유형')['피추천인매출'].sum().reset_index()
        fig=px.pie(ts_ref,values='피추천인매출',names='유형',hole=0.5,color='유형',color_discrete_map=tc)
        fig.update_traces(hovertemplate='%{label}<br>매출: %{customdata}<br>비중: %{percent}<extra></extra>',
                          customdata=[fmt_krw(v) for v in ts_ref['피추천인매출']])
        fig.update_layout(height=400,title=dict(text='추천인 유형별 피추천인 매출',x=0.01,font=dict(size=15)),
                          margin=dict(l=20,r=20,t=70,b=20))
        st.plotly_chart(fig, use_container_width=True)
    
    # 추천인별 테이블 (GAS: recommenderTable)
    st.markdown("##### 추천인별 현황")
    rtf=st.selectbox("추천인 유형 필터",["전체","영업팀","대리점","케어포"],key="ref_type")
    dr=rdf.copy()
    if rtf!="전체": dr=dr[dr['유형']==rtf]
    dr=dr.sort_values('피추천인매출',ascending=False)
    sr=st.text_input("🔍 추천인 검색",key="ref_search")
    if sr: dr=dr[dr.apply(lambda r:sr.lower() in str(r).lower(),axis=1)]
    st.dataframe(dr.style.format({'피추천인수':'{:,.0f}','피추천인매출':'{:,.0f}원'}),
                 use_container_width=True,height=500)

# ============================================================
# Tab 6. 케어포 멤버십 (GAS careforTab() 동일)
# ============================================================
with tab6:
    cfg=['케어포-시설','케어포-공생','케어포-주야간','케어포-방문','케어포-일반','케어포-종사자','케어포-보호자']
    co=filtered[filtered['회원 등급'].isin(cfg)]
    cmb=members[members['회원타입']=='케어포']
    cf_filtered=filtered_members[filtered_members['회원타입']=='케어포']
    
    cgf=st.selectbox("케어포 등급",["전체"]+cfg,key="cf_grade")
    if cgf!="전체":
        co=co[co['회원 등급']==cgf]
        cf_filtered=cf_filtered[cf_filtered['회원등급']==cgf]
    
    # KPI
    cbo=co.groupby('주문자 ID')['주문 ID'].nunique()
    crp=(cbo>=2).sum(); crr=crp/len(cbo)*100 if len(cbo)>0 else 0
    
    cols=st.columns(4)
    for col,(l,v,u) in zip(cols,[
        ("케어포 총 회원",fmt_num(len(cmb)),"처"),
        ("케어포 신규가입",fmt_num(len(cf_filtered)),"처"),
        ("케어포 주문회원",fmt_num(co['주문자 ID'].nunique()),"처"),
        ("케어포 재구매율",fmt_pct(crr),""),
    ]):
        col.markdown(kpi_card(l,v,u),unsafe_allow_html=True)
    
    # 등급별 매출/주문 (GAS: salesByGrade)
    cl,cr=st.columns(2)
    with cl:
        cga=co.groupby('회원 등급').agg(매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique'),주문회원수=('주문자 ID','nunique')).reset_index()
        cga['등급']=cga['회원 등급'].str.replace('케어포-','')
        fig=make_subplots(specs=[[{"secondary_y":True}]])
        fig.add_trace(go.Bar(x=cga['등급'],y=cga['매출'],name='매출액',marker_color='#3366CC',opacity=0.8,
                             hovertemplate='%{x}<br>매출: %{customdata}<extra></extra>',
                             customdata=[fmt_krw(v) for v in cga['매출']]),secondary_y=False)
        fig.add_trace(go.Scatter(x=cga['등급'],y=cga['주문건수'],name='주문건수',
                                 line=dict(color='#E8853D',width=2.5),mode='lines+markers',
                                 hovertemplate='%{x}<br>주문: %{y:,}건<extra></extra>'),secondary_y=True)
        fig.update_layout(height=420,title=dict(text='케어포 등급별 매출 · 주문',x=0.01,font=dict(size=15)),
                          margin=dict(l=60,r=50,t=70,b=30),
                          legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=10)))
        st.plotly_chart(fig, use_container_width=True)
    with cr:
        # 전용 상품 매출 추이 (GAS: productTrend)
        cpd=co[co['상품명'].str.contains(r'\[케어포',na=False)]
        cpm=cpd.groupby('주문월')['판매합계금액'].sum().reset_index()
        cpm['주문월_kr'] = ym_series_kr(cpm['주문월'])
        fig=px.area(cpm,x='주문월_kr',y='판매합계금액',color_discrete_sequence=['#27AE60'])
        fig.update_traces(hovertemplate='%{x}<br>매출: %{customdata}<extra></extra>',
                          customdata=[fmt_krw(v) for v in cpm['판매합계금액']])
        fig.update_layout(height=420,title=dict(text='케어포 전용 상품 매출 추이',x=0.01,font=dict(size=15)),
                          margin=dict(l=60,r=30,t=70,b=50),showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    # 신규가입 추이 (GAS: newCfMembers - 필터된 케어포 회원 기준)
    cj=cf_filtered.groupby(['가입월','회원등급']).size().reset_index(name='가입자수')
    cj['가입월_kr'] = ym_series_kr(cj['가입월'])
    ccl={'케어포-시설':'#3366CC','케어포-공생':'#E8853D','케어포-주야간':'#27AE60',
         '케어포-방문':'#9B59B6','케어포-일반':'#F39C12','케어포-종사자':'#E74C3C','케어포-보호자':'#1ABC9C'}
    fig=px.bar(cj,x='가입월_kr',y='가입자수',color='회원등급',color_discrete_map=ccl)
    for tr in fig.data:
        tr.hovertemplate = '%{x}<br>' + tr.name + ': %{y:,}처<extra></extra>'
    fig.update_layout(height=420,barmode='stack',
                      title=dict(text='케어포 등급별 신규가입 추이',x=0.01,font=dict(size=15)),
                      margin=dict(l=50,r=20,t=70,b=80),
                      legend=dict(orientation="h",yanchor="top",y=-0.15,x=0,font=dict(size=9)))
    st.plotly_chart(fig, use_container_width=True)

# 푸터
st.markdown("---")
st.markdown(f"<p style='text-align:center;color:#94a3b8;font-size:0.8rem;'>© 대상웰라이프 B2B몰 대시보드 · 데이터 기준: {pd.Timestamp.now().strftime('%Y년 %m월 %d일')}</p>",unsafe_allow_html=True)
