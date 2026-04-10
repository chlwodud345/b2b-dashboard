import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.io as pio

# ============================================================
# 페이지 설정
# ============================================================
st.set_page_config(
    page_title="대상웰라이프 B2B몰 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# Plotly 테마
# ============================================================
COLORS = ['#3366CC','#E8853D','#27AE60','#9B59B6','#E74C3C',
          '#1ABC9C','#F39C12','#2980B9','#8E44AD','#D35400']

CHART_TEMPLATE = go.layout.Template(
    layout=go.Layout(
        font=dict(family="Noto Sans KR, sans-serif", size=13, color="#1e293b"),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=50, r=20, t=40, b=50),
        xaxis=dict(gridcolor="#f1f5f9", showline=True, linecolor="#e2e8f0"),
        yaxis=dict(gridcolor="#f1f5f9", showline=True, linecolor="#e2e8f0"),
        colorway=COLORS,
    )
)
pio.templates["dashboard"] = CHART_TEMPLATE
pio.templates.default = "dashboard"

# ============================================================
# 커스텀 CSS
# ============================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="st-"] { font-family: 'Noto Sans KR', sans-serif; }
    
    /* 헤더 */
    .main-header {
        background: linear-gradient(135deg, #1B2A4A 0%, #2D4A7A 100%);
        color: white;
        padding: 20px 30px;
        border-radius: 12px;
        margin-bottom: 24px;
    }
    .main-header h1 { margin: 0; font-size: 1.5rem; font-weight: 700; }
    .main-header p { margin: 4px 0 0; opacity: 0.7; font-size: 0.85rem; }
    
    /* KPI 카드 */
    .kpi-card {
        background: white;
        border-radius: 12px;
        padding: 20px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        text-align: center;
        transition: box-shadow 0.2s;
    }
    .kpi-card:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.08); }
    .kpi-label { font-size: 0.8rem; color: #64748b; font-weight: 500; margin-bottom: 4px; }
    .kpi-value { font-size: 1.6rem; font-weight: 700; color: #1e293b; }
    .kpi-unit { font-size: 0.85rem; color: #94a3b8; margin-left: 2px; }
    
    /* 탭 스타일 */
    .stTabs [data-baseweb="tab-list"] { gap: 4px; }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 20px;
        font-weight: 500;
    }
    
    /* 데이터프레임 */
    .stDataFrame { border-radius: 8px; }
    
    /* 사이드바 */
    [data-testid="stSidebar"] { background: #f8fafc; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 포맷팅 함수
# ============================================================
def fmt_currency(n):
    if pd.isna(n) or n == 0:
        return "0원"
    if abs(n) >= 1e8:
        return f"{n/1e8:.1f}억"
    if abs(n) >= 1e4:
        return f"{n/1e4:,.0f}만"
    return f"{n:,.0f}원"

def fmt_number(n):
    if pd.isna(n):
        return "0"
    return f"{n:,.0f}"

def fmt_percent(n):
    if pd.isna(n):
        return "0%"
    return f"{n:.1f}%"

def kpi_card(label, value, unit=""):
    return f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}<span class="kpi-unit">{unit}</span></div>
    </div>
    """

# ============================================================
# 데이터 전처리 함수
# ============================================================
@st.cache_data
def process_data(orders, members, referrals_df):
    # 주문내역 전처리
    orders['주문일'] = pd.to_datetime(orders['주문일'], errors='coerce')
    orders['주문일자'] = orders['주문일'].dt.strftime('%Y-%m-%d')
    orders['주문월'] = orders['주문일'].dt.to_period('M').astype(str)
    orders['주문시간'] = orders['주문일'].dt.hour
    dow_map = {'Monday':'월','Tuesday':'화','Wednesday':'수','Thursday':'목',
               'Friday':'금','Saturday':'토','Sunday':'일'}
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
        for r in ['서울','부산','대구','인천','광주','대전','울산','세종',
                   '경기','강원','충북','충남','전북','전남','경북','경남','제주']:
            if addr.startswith(r): return r
        return '기타'
    
    orders['지역'] = orders['주소(주문자)'].apply(extract_region)
    
    # 회원정보 전처리
    members['가입일'] = pd.to_datetime(members['가입일'], errors='coerce')
    members['가입월'] = members['가입일'].dt.to_period('M').astype(str)
    members['사업자번호'] = members['사업자번호'].astype(str).str.replace('-','').str.strip()
    
    # 추천인 전처리
    referrals_df['피추천인 사업자 번호'] = referrals_df['피추천인 사업자 번호'].astype(str).str.replace('-','').str.strip()
    
    return orders, members, referrals_df

# ============================================================
# 데이터 로드 (구글 드라이브에서 직접 다운로드)
# ============================================================
import io
import requests

GDRIVE_FILE_ID = '1Op9Y2FFb_aLQJKAcLyKj9HJQbK6YYnmf'

def download_from_gdrive(file_id):
    """구글 드라이브 대용량 파일 다운로드 (바이러스 검사 우회)"""
    session = requests.Session()
    url = f'https://drive.google.com/uc?export=download&id={file_id}'
    
    response = session.get(url, stream=True)
    
    # 대용량 파일: 바이러스 검사 확인 페이지 우회
    confirm_token = None
    for key, value in response.cookies.items():
        if key.startswith('download_warning'):
            confirm_token = value
            break
    
    if confirm_token:
        url = f'https://drive.google.com/uc?export=download&confirm={confirm_token}&id={file_id}'
        response = session.get(url, stream=True)
    
    # 그래도 HTML이 반환되면 confirm=t 방식 시도
    content = response.content
    if content[:4] != b'PK\x03\x04':  # xlsx는 ZIP 형식이라 PK로 시작
        url = f'https://drive.google.com/uc?export=download&confirm=t&id={file_id}'
        response = session.get(url, stream=True)
        content = response.content
    
    return io.BytesIO(content)

@st.cache_data(ttl=3600, show_spinner="📥 구글 드라이브에서 데이터를 불러오는 중...")
def load_from_gdrive():
    file_bytes = download_from_gdrive(GDRIVE_FILE_ID)
    
    orders_raw = pd.read_excel(file_bytes, sheet_name='주문내역', header=1, engine='openpyxl')
    file_bytes.seek(0)
    members_raw = pd.read_excel(file_bytes, sheet_name='회원정보', header=1, engine='openpyxl')
    file_bytes.seek(0)
    referrals_raw = pd.read_excel(file_bytes, sheet_name='추천인', header=1, engine='openpyxl')
    
    return process_data(orders_raw, members_raw, referrals_raw)

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

# ============================================================
# 사이드바 필터
# ============================================================
st.sidebar.markdown("## 🔍 필터")

# 연도/월
years = sorted(orders['주문일'].dt.year.dropna().unique().astype(int))
selected_year = st.sidebar.selectbox("연도", ["전체"] + [str(y) for y in years], index=0)

if selected_year != "전체":
    month_options = sorted(orders[orders['주문일'].dt.year == int(selected_year)]['주문일'].dt.month.dropna().unique().astype(int))
    selected_month = st.sidebar.selectbox("월", ["전체"] + [f"{m}월" for m in month_options], index=0)
else:
    selected_month = "전체"

# 회원구분 / 회원등급
member_types = ["전체"] + sorted(orders['주문자 구분'].dropna().unique().tolist())
selected_type = st.sidebar.selectbox("회원구분", member_types, index=0)

member_grades = ["전체"] + sorted(orders['회원 등급'].dropna().unique().tolist())
selected_grade = st.sidebar.selectbox("회원등급", member_grades, index=0)

# ---- 필터 적용 ----
filtered = orders.copy()
if selected_year != "전체":
    filtered = filtered[filtered['주문일'].dt.year == int(selected_year)]
if selected_month != "전체":
    m = int(selected_month.replace('월',''))
    filtered = filtered[filtered['주문일'].dt.month == m]
if selected_type != "전체":
    filtered = filtered[filtered['주문자 구분'] == selected_type]
if selected_grade != "전체":
    filtered = filtered[filtered['회원 등급'] == selected_grade]

# ============================================================
# 헤더
# ============================================================
st.markdown("""
<div class="main-header">
    <h1>📊 대상웰라이프 B2B몰 대시보드</h1>
    <p>Sales & Operations Analytics</p>
</div>
""", unsafe_allow_html=True)

# ============================================================
# 탭 구성
# ============================================================
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📋 종합 현황", "💰 매출 분석", "📦 상품 분석",
    "👥 회원 분석", "🔗 추천인 분석", "💚 케어포 멤버십"
])

# ============================================================
# Tab 1. 종합 현황
# ============================================================
with tab1:
    total_sales = filtered['판매합계금액'].sum()
    total_orders = filtered['주문 ID'].nunique()
    total_buyers = filtered['주문자 ID'].nunique()
    total_members = len(members)
    avg_order = total_sales / total_orders if total_orders > 0 else 0
    
    cols = st.columns(5)
    kpis = [
        ("총 매출액", fmt_currency(total_sales), ""),
        ("총 주문건수", fmt_number(total_orders), "건"),
        ("총 회원수", fmt_number(total_members), "처"),
        ("주문회원수", fmt_number(total_buyers), "처"),
        ("객단가", fmt_currency(avg_order), ""),
    ]
    for col, (label, value, unit) in zip(cols, kpis):
        col.markdown(kpi_card(label, value, unit), unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 월별 매출 · 주문건수 추이
    col_left, col_right = st.columns(2)
    
    with col_left:
        monthly = filtered.groupby('주문월').agg(
            매출=('판매합계금액', 'sum'),
            주문건수=('주문 ID', 'nunique')
        ).reset_index()
        
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(
            go.Bar(x=monthly['주문월'], y=monthly['매출'], name='매출액',
                   marker_color='#3366CC', opacity=0.8),
            secondary_y=False
        )
        fig.add_trace(
            go.Scatter(x=monthly['주문월'], y=monthly['주문건수'], name='주문건수',
                       line=dict(color='#E8853D', width=2.5), mode='lines+markers'),
            secondary_y=True
        )
        fig.update_layout(title='월별 매출 · 주문건수 추이', height=400,
                          legend=dict(orientation="h", y=1.12))
        fig.update_yaxes(title_text="매출액 (원)", secondary_y=False)
        fig.update_yaxes(title_text="주문건수", secondary_y=True)
        st.plotly_chart(fig, use_container_width=True)
    
    with col_right:
        type_sales = filtered.groupby('주문자 구분')['판매합계금액'].sum().reset_index()
        type_sales.columns = ['구분', '매출']
        fig = px.pie(type_sales, values='매출', names='구분', hole=0.5,
                     color_discrete_sequence=COLORS)
        fig.update_layout(title='회원구분별 매출 비중', height=400)
        fig.update_traces(textinfo='percent+label', textfont_size=11)
        st.plotly_chart(fig, use_container_width=True)
    
    # 일별 매출 + 지역별 매출
    col_left, col_right = st.columns(2)
    
    with col_left:
        daily = filtered.groupby('주문일자')['판매합계금액'].sum().reset_index()
        daily.columns = ['날짜', '매출']
        fig = px.area(daily, x='날짜', y='매출', color_discrete_sequence=['#3366CC'])
        fig.update_layout(title='일별 매출 추이', height=350)
        st.plotly_chart(fig, use_container_width=True)
    
    with col_right:
        region = filtered.groupby('지역')['판매합계금액'].sum().sort_values().reset_index()
        region.columns = ['지역', '매출']
        fig = px.bar(region, x='매출', y='지역', orientation='h',
                     color_discrete_sequence=COLORS)
        fig.update_layout(title='지역별 매출', height=350)
        st.plotly_chart(fig, use_container_width=True)

# ============================================================
# Tab 2. 매출 분석
# ============================================================
with tab2:
    # 회원구분별 × 월별 매출
    col_left, col_right = st.columns(2)
    
    with col_left:
        type_month = filtered.groupby(['주문월', '주문자 구분'])['판매합계금액'].sum().reset_index()
        fig = px.bar(type_month, x='주문월', y='판매합계금액', color='주문자 구분',
                     color_discrete_sequence=COLORS)
        fig.update_layout(title='회원구분별 × 월별 매출 추이', height=400,
                          barmode='stack', legend=dict(orientation="h", y=1.12))
        st.plotly_chart(fig, use_container_width=True)
    
    with col_right:
        grade_sales = filtered.groupby('회원 등급')['판매합계금액'].sum().sort_values().reset_index()
        fig = px.bar(grade_sales, x='판매합계금액', y='회원 등급', orientation='h',
                     color_discrete_sequence=COLORS)
        fig.update_layout(title='회원등급별 매출', height=400)
        st.plotly_chart(fig, use_container_width=True)
    
    # 요일 · 시간대 히트맵
    st.subheader("요일 · 시간대별 주문 매출 히트맵")
    dow_order = ['월','화','수','목','금','토','일']
    heatmap_data = filtered.groupby(['요일','주문시간'])['판매합계금액'].sum().reset_index()
    heatmap_pivot = heatmap_data.pivot_table(index='요일', columns='주문시간', values='판매합계금액', fill_value=0)
    heatmap_pivot = heatmap_pivot.reindex(dow_order)
    
    fig = go.Figure(data=go.Heatmap(
        z=heatmap_pivot.values,
        x=[f'{h}시' for h in heatmap_pivot.columns],
        y=heatmap_pivot.index,
        colorscale=[[0,'#F0F2F5'],[0.5,'#6B9BD2'],[1,'#1B2A4A']],
        hovertemplate='%{y} %{x}: %{z:,.0f}원<extra></extra>'
    ))
    fig.update_layout(height=300, margin=dict(l=40, r=20, t=20, b=40))
    st.plotly_chart(fig, use_container_width=True)
    
    # 기관별 매출 테이블
    st.subheader("기관별 매출 현황")
    
    buyer_agg = filtered.groupby(['주문자 ID','주문자명','주문자 구분','회원 등급']).agg(
        매출=('판매합계금액', 'sum'),
        주문건수=('주문 ID', 'nunique'),
        최근주문일=('주문일자', 'max')
    ).reset_index()
    buyer_agg['객단가'] = (buyer_agg['매출'] / buyer_agg['주문건수']).round(0)
    buyer_agg = buyer_agg.sort_values('매출', ascending=False)
    
    search = st.text_input("🔍 검색 (아이디, 주문자명, 회원구분)", key="sales_search")
    if search:
        mask = buyer_agg.apply(lambda row: search.lower() in str(row).lower(), axis=1)
        buyer_agg = buyer_agg[mask]
    
    st.dataframe(
        buyer_agg.style.format({
            '매출': '{:,.0f}원',
            '주문건수': '{:,.0f}건',
            '객단가': '{:,.0f}원'
        }),
        use_container_width=True,
        height=500
    )

# ============================================================
# Tab 3. 상품 분석
# ============================================================
with tab3:
    # 파레토 차트
    product_agg = filtered.groupby(['상품명','상품 코드']).agg(
        매출=('판매합계금액', 'sum'),
        수량=('주문 수량', 'sum'),
        주문건수=('주문 ID', 'nunique')
    ).reset_index().sort_values('매출', ascending=False)
    
    top20 = product_agg.head(20).copy()
    total_sales_all = product_agg['매출'].sum()
    top20['누적비중'] = (top20['매출'].cumsum() / total_sales_all * 100).round(1)
    
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(
        go.Bar(x=top20['상품명'].str[:20], y=top20['매출'], name='매출액',
               marker_color='#3366CC', opacity=0.8),
        secondary_y=False
    )
    fig.add_trace(
        go.Scatter(x=top20['상품명'].str[:20], y=top20['누적비중'], name='누적 비중',
                   line=dict(color='#E8853D', width=2.5), mode='lines+markers'),
        secondary_y=True
    )
    fig.update_layout(title='상품별 매출 TOP 20 (파레토 차트)', height=450,
                      legend=dict(orientation="h", y=1.12),
                      xaxis=dict(tickangle=45, tickfont=dict(size=9)))
    fig.update_yaxes(title_text="매출액 (원)", secondary_y=False)
    fig.update_yaxes(title_text="누적 비중 (%)", range=[0, 105], secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)
    
    # 전체 상품 테이블
    st.subheader("전체 상품 매출 현황")
    search_p = st.text_input("🔍 상품명/상품코드 검색", key="product_search")
    display_products = product_agg.copy()
    if search_p:
        mask = display_products.apply(lambda row: search_p.lower() in str(row).lower(), axis=1)
        display_products = display_products[mask]
    
    st.dataframe(
        display_products.style.format({
            '매출': '{:,.0f}원',
            '수량': '{:,.0f}',
            '주문건수': '{:,.0f}건'
        }),
        use_container_width=True,
        height=400
    )
    
    # 회원구분별 × 상품 크로스
    st.subheader("회원구분별 × 상품 매출 크로스 (TOP 20)")
    top20_names = product_agg.head(20)['상품명'].tolist()
    cross_data = filtered[filtered['상품명'].isin(top20_names)]
    cross_pivot = cross_data.pivot_table(
        index='상품명', columns='주문자 구분', values='판매합계금액',
        aggfunc='sum', fill_value=0
    )
    cross_pivot['합계'] = cross_pivot.sum(axis=1)
    cross_pivot = cross_pivot.sort_values('합계', ascending=False)
    
    st.dataframe(
        cross_pivot.style.format('{:,.0f}원'),
        use_container_width=True,
        height=500
    )
    
    # 월별 상품 매출 추이
    st.subheader("월별 상품 매출 추이")
    top5_names = product_agg.head(5)['상품명'].tolist()
    selected_products = st.multiselect("상품 선택", product_agg['상품명'].tolist(),
                                        default=top5_names, key="product_trend")
    
    if selected_products:
        trend_data = filtered[filtered['상품명'].isin(selected_products)]
        trend_monthly = trend_data.groupby(['주문월','상품명'])['판매합계금액'].sum().reset_index()
        fig = px.line(trend_monthly, x='주문월', y='판매합계금액', color='상품명',
                      markers=True, color_discrete_sequence=COLORS)
        fig.update_layout(title='월별 상품 매출 추이', height=400,
                          legend=dict(orientation="h", y=-0.2))
        for trace in fig.data:
            if len(trace.name) > 25:
                trace.name = trace.name[:25] + '...'
        st.plotly_chart(fig, use_container_width=True)

# ============================================================
# Tab 4. 회원 분석
# ============================================================
with tab4:
    # 회원별 주문 집계
    member_orders = orders.groupby('주문자 ID').agg(
        첫주문일=('주문일', 'min'),
        주문건수=('주문 ID', 'nunique'),
        주문월수=('주문월', 'nunique')
    ).reset_index()
    
    # KPI
    converted = members[members['아이디'].isin(orders['주문자 ID'].unique())]
    conversion_rate = len(converted) / len(members) * 100 if len(members) > 0 else 0
    
    repeat_buyers = member_orders[member_orders['주문건수'] >= 2]
    repeat_rate = len(repeat_buyers) / len(member_orders) * 100 if len(member_orders) > 0 else 0
    
    recent_3m = orders[orders['주문일'] >= orders['주문일'].max() - pd.DateOffset(months=3)]
    active_count = recent_3m['주문자 ID'].nunique()
    
    cols = st.columns(5)
    kpis = [
        ("총 회원수", fmt_number(len(members)), "처"),
        ("구매전환율", fmt_percent(conversion_rate), ""),
        ("재구매율", fmt_percent(repeat_rate), ""),
        ("활성회원", fmt_number(active_count), "처"),
        ("총 주문회원", fmt_number(len(member_orders)), "처"),
    ]
    for col, (label, value, unit) in zip(cols, kpis):
        col.markdown(kpi_card(label, value, unit), unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 월별 신규가입 + 등급별 분포
    col_left, col_right = st.columns(2)
    
    with col_left:
        join_monthly = members.groupby(['가입월','회원타입']).size().reset_index(name='가입자수')
        fig = px.bar(join_monthly, x='가입월', y='가입자수', color='회원타입',
                     color_discrete_sequence=COLORS)
        fig.update_layout(title='월별 신규가입자 추이 (회원타입별)', height=380,
                          barmode='stack', legend=dict(orientation="h", y=1.12))
        st.plotly_chart(fig, use_container_width=True)
    
    with col_right:
        grade_dist = members['회원등급'].value_counts().reset_index()
        grade_dist.columns = ['등급', '수']
        fig = px.bar(grade_dist, x='수', y='등급', orientation='h',
                     color_discrete_sequence=COLORS)
        fig.update_layout(title='회원등급별 가입자 분포', height=380)
        st.plotly_chart(fig, use_container_width=True)
    
    # 첫 주문 소요일 + 주문횟수 분포
    col_left, col_right = st.columns(2)
    
    with col_left:
        merged = members.merge(member_orders, left_on='아이디', right_on='주문자 ID', how='inner')
        merged['소요일'] = (merged['첫주문일'] - merged['가입일']).dt.days
        merged = merged[merged['소요일'] >= 0]
        
        bins = [0, 1, 8, 15, 31, 61, 91, 9999]
        labels = ['당일','1~7일','8~14일','15~30일','31~60일','61~90일','90일+']
        merged['소요구간'] = pd.cut(merged['소요일'], bins=bins, labels=labels, right=False)
        days_hist = merged['소요구간'].value_counts().reindex(labels).fillna(0).reset_index()
        days_hist.columns = ['구간', '회원수']
        
        fig = px.bar(days_hist, x='구간', y='회원수', color_discrete_sequence=['#3366CC'])
        fig.update_layout(title='가입 후 첫 주문까지 소요일', height=350)
        st.plotly_chart(fig, use_container_width=True)
    
    with col_right:
        bins_oc = [1, 2, 4, 6, 11, 21, 9999]
        labels_oc = ['1회','2~3회','4~5회','6~10회','11~20회','20회+']
        member_orders['주문구간'] = pd.cut(member_orders['주문건수'], bins=bins_oc, labels=labels_oc, right=False)
        oc_dist = member_orders['주문구간'].value_counts().reindex(labels_oc).fillna(0).reset_index()
        oc_dist.columns = ['구간', '회원수']
        
        fig = px.bar(oc_dist, x='구간', y='회원수', color_discrete_sequence=COLORS)
        fig.update_layout(title='주문횟수 구간별 회원 분포', height=350)
        st.plotly_chart(fig, use_container_width=True)
    
    # 코호트 리텐션
    st.subheader("코호트 리텐션 히트맵")
    
    cohort_members = members[members['가입일'].notna()].copy()
    cohort_members['코호트'] = cohort_members['가입일'].dt.to_period('M').astype(str)
    
    order_months = orders.groupby('주문자 ID')['주문월'].apply(set).to_dict()
    
    cohorts = sorted(cohort_members['코호트'].unique())
    max_offset = 12
    retention_data = []
    
    for cohort in cohorts:
        cohort_users = cohort_members[cohort_members['코호트'] == cohort]['아이디'].tolist()
        size = len(cohort_users)
        if size == 0:
            continue
        
        row = {'코호트': cohort, '코호트크기': size}
        cohort_period = pd.Period(cohort, freq='M')
        
        for offset in range(max_offset):
            target = str(cohort_period + offset)
            active = sum(1 for uid in cohort_users if target in order_months.get(uid, set()))
            row[f'{offset}개월'] = round(active / size * 100, 1) if size > 0 else 0
        
        retention_data.append(row)
    
    if retention_data:
        ret_df = pd.DataFrame(retention_data)
        month_cols = [f'{i}개월' for i in range(max_offset)]
        
        z_vals = ret_df[month_cols].values
        
        fig = go.Figure(data=go.Heatmap(
            z=z_vals,
            x=month_cols,
            y=[f"{r['코호트']} ({r['코호트크기']}처)" for _, r in ret_df.iterrows()],
            colorscale=[[0,'#F0F2F5'],[0.3,'#A8D5A2'],[1,'#27AE60']],
            text=[[f'{v:.1f}%' if v > 0 else '-' for v in row] for row in z_vals],
            texttemplate='%{text}',
            textfont=dict(size=10),
            hovertemplate='%{y}<br>%{x}: %{z:.1f}%<extra></extra>'
        ))
        fig.update_layout(height=max(300, len(retention_data) * 28 + 100),
                          margin=dict(l=150, r=20, t=20, b=40),
                          yaxis=dict(tickfont=dict(size=10)))
        st.plotly_chart(fig, use_container_width=True)

# ============================================================
# Tab 5. 추천인 분석
# ============================================================
with tab5:
    # 추천인 분류
    class_map = {}
    for _, r in referrals_df.iterrows():
        name = str(r.get('추천인', '')).strip()
        group = str(r.get('회원그룹', ''))
        if not name or name == '-' or name == 'nan':
            continue
        if group == '영업팀':
            class_map[name] = '케어포' if name == '케어포' else '영업팀'
        elif group == '대리점 회원':
            class_map[name] = '대리점'
    
    rec_agg = {}
    for _, r in referrals_df.iterrows():
        name = str(r.get('추천인', '')).strip()
        if not name or name == '-' or name == 'nan':
            continue
        if name not in rec_agg:
            rec_agg[name] = {
                '추천인': name,
                '유형': class_map.get(name, '케어포'),
                '추천인코드': r.get('추천인코드', ''),
                '피추천인수': 0,
                '사업자번호목록': []
            }
        biz = str(r.get('피추천인 사업자 번호', '')).strip()
        if biz and biz != '-' and biz != 'nan':
            rec_agg[name]['피추천인수'] += 1
            rec_agg[name]['사업자번호목록'].append(biz)
    
    # 피추천인 매출 연결
    biz_to_uid = members.set_index('사업자번호')['아이디'].to_dict()
    buyer_sales = filtered.groupby('주문자 ID')['판매합계금액'].sum().to_dict()
    
    for name in rec_agg:
        total = 0
        for biz in rec_agg[name]['사업자번호목록']:
            uid = biz_to_uid.get(biz)
            if uid and uid in buyer_sales:
                total += buyer_sales[uid]
        rec_agg[name]['피추천인매출'] = total
    
    rec_df = pd.DataFrame(rec_agg.values())[['추천인','유형','추천인코드','피추천인수','피추천인매출']]
    
    # KPI
    cols = st.columns(3)
    kpis = [
        ("총 추천인 수", fmt_number(len(rec_df)), "회원"),
        ("총 피추천인 수", fmt_number(rec_df['피추천인수'].sum()), "회원"),
        ("추천인당 평균 피추천인", f"{rec_df['피추천인수'].mean():.1f}" if len(rec_df) > 0 else "0", "회원"),
    ]
    for col, (label, value, unit) in zip(cols, kpis):
        col.markdown(kpi_card(label, value, unit), unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 유형별 피추천인 수 + 매출 비중
    col_left, col_right = st.columns(2)
    
    with col_left:
        type_ref = rec_df.groupby('유형')['피추천인수'].sum().reset_index()
        type_colors = {'영업팀':'#3366CC','대리점':'#E8853D','케어포':'#27AE60'}
        fig = px.bar(type_ref, x='유형', y='피추천인수',
                     color='유형', color_discrete_map=type_colors)
        fig.update_layout(title='추천인 유형별 피추천인 수', height=350, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    with col_right:
        type_sales_ref = rec_df.groupby('유형')['피추천인매출'].sum().reset_index()
        fig = px.pie(type_sales_ref, values='피추천인매출', names='유형',
                     hole=0.5, color='유형', color_discrete_map=type_colors)
        fig.update_layout(title='추천인 유형별 피추천인 매출', height=350)
        st.plotly_chart(fig, use_container_width=True)
    
    # 추천인별 테이블
    st.subheader("추천인별 현황")
    
    ref_type_filter = st.selectbox("추천인 유형 필터", ["전체","영업팀","대리점","케어포"], key="ref_type")
    display_ref = rec_df.copy()
    if ref_type_filter != "전체":
        display_ref = display_ref[display_ref['유형'] == ref_type_filter]
    
    display_ref = display_ref.sort_values('피추천인매출', ascending=False)
    
    search_ref = st.text_input("🔍 추천인 검색", key="ref_search")
    if search_ref:
        mask = display_ref.apply(lambda row: search_ref.lower() in str(row).lower(), axis=1)
        display_ref = display_ref[mask]
    
    st.dataframe(
        display_ref.style.format({
            '피추천인수': '{:,.0f}',
            '피추천인매출': '{:,.0f}원'
        }),
        use_container_width=True,
        height=500
    )

# ============================================================
# Tab 6. 케어포 멤버십
# ============================================================
with tab6:
    cf_grades = ['케어포-시설','케어포-공생','케어포-주야간','케어포-방문',
                 '케어포-일반','케어포-종사자','케어포-보호자']
    
    cf_orders = filtered[filtered['회원 등급'].isin(cf_grades)]
    cf_members = members[members['회원타입'] == '케어포']
    
    # 케어포 등급 필터
    cf_grade_filter = st.selectbox("케어포 등급", ["전체"] + cf_grades, key="cf_grade")
    if cf_grade_filter != "전체":
        cf_orders = cf_orders[cf_orders['회원 등급'] == cf_grade_filter]
    
    # KPI
    cf_buyer_orders = cf_orders.groupby('주문자 ID')['주문 ID'].nunique()
    cf_repeat = (cf_buyer_orders >= 2).sum()
    cf_repeat_rate = cf_repeat / len(cf_buyer_orders) * 100 if len(cf_buyer_orders) > 0 else 0
    
    cols = st.columns(3)
    kpis = [
        ("케어포 총 회원", fmt_number(len(cf_members)), "처"),
        ("케어포 주문회원", fmt_number(cf_orders['주문자 ID'].nunique()), "처"),
        ("케어포 재구매율", fmt_percent(cf_repeat_rate), ""),
    ]
    for col, (label, value, unit) in zip(cols, kpis):
        col.markdown(kpi_card(label, value, unit), unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 등급별 매출/주문 + 전용 상품 추이
    col_left, col_right = st.columns(2)
    
    with col_left:
        cf_grade_agg = cf_orders.groupby('회원 등급').agg(
            매출=('판매합계금액', 'sum'),
            주문건수=('주문 ID', 'nunique')
        ).reset_index()
        cf_grade_agg['등급'] = cf_grade_agg['회원 등급'].str.replace('케어포-','')
        
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(
            go.Bar(x=cf_grade_agg['등급'], y=cf_grade_agg['매출'], name='매출액',
                   marker_color='#3366CC', opacity=0.8),
            secondary_y=False
        )
        fig.add_trace(
            go.Scatter(x=cf_grade_agg['등급'], y=cf_grade_agg['주문건수'], name='주문건수',
                       line=dict(color='#E8853D', width=2.5), mode='lines+markers'),
            secondary_y=True
        )
        fig.update_layout(title='케어포 등급별 매출 · 주문', height=400,
                          legend=dict(orientation="h", y=1.12))
        st.plotly_chart(fig, use_container_width=True)
    
    with col_right:
        cf_product = cf_orders[cf_orders['상품명'].str.contains(r'\[케어포', na=False)]
        cf_product_monthly = cf_product.groupby('주문월')['판매합계금액'].sum().reset_index()
        fig = px.area(cf_product_monthly, x='주문월', y='판매합계금액',
                      color_discrete_sequence=['#27AE60'])
        fig.update_layout(title='케어포 전용 상품 매출 추이', height=400)
        st.plotly_chart(fig, use_container_width=True)
    
    # 신규가입 추이
    cf_join = cf_members.groupby(['가입월','회원등급']).size().reset_index(name='가입자수')
    cf_colors = {
        '케어포-시설':'#3366CC','케어포-공생':'#E8853D','케어포-주야간':'#27AE60',
        '케어포-방문':'#9B59B6','케어포-일반':'#F39C12','케어포-종사자':'#E74C3C',
        '케어포-보호자':'#1ABC9C'
    }
    fig = px.bar(cf_join, x='가입월', y='가입자수', color='회원등급',
                 color_discrete_map=cf_colors)
    fig.update_layout(title='케어포 등급별 신규가입 추이', height=380,
                      barmode='stack', legend=dict(orientation="h", y=-0.15))
    st.plotly_chart(fig, use_container_width=True)

# ============================================================
# 푸터
# ============================================================
st.markdown("---")
st.markdown(
    f"<p style='text-align:center;color:#94a3b8;font-size:0.8rem;'>"
    f"© 대상웰라이프 B2B몰 대시보드 · 데이터 기준: {pd.Timestamp.now().strftime('%Y-%m-%d')}</p>",
    unsafe_allow_html=True
)
