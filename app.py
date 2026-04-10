import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.io as pio
import io
import requests

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

# ============================================================
# 커스텀 CSS
# ============================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700&display=swap');
    html, body, [class*="st-"] { font-family: 'Noto Sans KR', sans-serif; }
    .main-header {
        background: linear-gradient(135deg, #1B2A4A 0%, #2D4A7A 100%);
        color: white; padding: 20px 30px; border-radius: 12px; margin-bottom: 24px;
    }
    .main-header h1 { margin: 0; font-size: 1.5rem; font-weight: 700; }
    .main-header p { margin: 4px 0 0; opacity: 0.7; font-size: 0.85rem; }
    .kpi-card {
        background: white; border-radius: 12px; padding: 20px;
        border: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        text-align: center; margin-bottom: 16px;
    }
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
def fmt_currency(n):
    if pd.isna(n) or n == 0: return "0원"
    if abs(n) >= 1e8: return f"{n/1e8:.1f}억"
    if abs(n) >= 1e4: return f"{n/1e4:,.0f}만"
    return f"{n:,.0f}원"

def fmt_number(n):
    if pd.isna(n): return "0"
    return f"{n:,.0f}"

def fmt_percent(n):
    if pd.isna(n): return "0%"
    return f"{n:.1f}%"

def kpi_card(label, value, unit=""):
    return f'<div class="kpi-card"><div class="kpi-label">{label}</div><div class="kpi-value">{value}<span class="kpi-unit">{unit}</span></div></div>'

# ============================================================
# 데이터 전처리 함수
# ============================================================
@st.cache_data
def process_data(orders, members, referrals_df):
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
        if key.startswith('download_warning'):
            confirm_token = value
            break
    if confirm_token:
        url = f'https://drive.google.com/uc?export=download&confirm={confirm_token}&id={file_id}'
        response = session.get(url, stream=True)
    content = response.content
    if content[:4] != b'PK\x03\x04':
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
years = sorted(orders['주문일'].dt.year.dropna().unique().astype(int))
selected_year = st.sidebar.selectbox("연도", ["전체"] + [str(y) for y in years], index=0)
if selected_year != "전체":
    month_options = sorted(orders[orders['주문일'].dt.year == int(selected_year)]['주문일'].dt.month.dropna().unique().astype(int))
    selected_month = st.sidebar.selectbox("월", ["전체"] + [f"{m}월" for m in month_options], index=0)
else:
    selected_month = "전체"
member_types = ["전체"] + sorted(orders['주문자 구분'].dropna().unique().tolist())
selected_type = st.sidebar.selectbox("회원구분", member_types, index=0)
member_grades = ["전체"] + sorted(orders['회원 등급'].dropna().unique().tolist())
selected_grade = st.sidebar.selectbox("회원등급", member_grades, index=0)

filtered = orders.copy()
if selected_year != "전체":
    filtered = filtered[filtered['주문일'].dt.year == int(selected_year)]
if selected_month != "전체":
    filtered = filtered[filtered['주문일'].dt.month == int(selected_month.replace('월',''))]
if selected_type != "전체":
    filtered = filtered[filtered['주문자 구분'] == selected_type]
if selected_grade != "전체":
    filtered = filtered[filtered['회원 등급'] == selected_grade]

# ============================================================
# 헤더
# ============================================================
st.markdown('<div class="main-header"><h1>📊 대상웰라이프 B2B몰 대시보드</h1><p>Sales & Operations Analytics</p></div>', unsafe_allow_html=True)

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
    for col, (l, v, u) in zip(cols, [
        ("총 매출액", fmt_currency(total_sales), ""), ("총 주문건수", fmt_number(total_orders), "건"),
        ("총 회원수", fmt_number(total_members), "처"), ("주문회원수", fmt_number(total_buyers), "처"),
        ("객단가", fmt_currency(avg_order), ""),
    ]):
        col.markdown(kpi_card(l, v, u), unsafe_allow_html=True)
    
    # 월별 매출 · 주문건수
    monthly = filtered.groupby('주문월').agg(매출=('판매합계금액','sum'), 주문건수=('주문 ID','nunique')).reset_index()
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=monthly['주문월'], y=monthly['매출'], name='매출액', marker_color='#3366CC', opacity=0.8), secondary_y=False)
    fig.add_trace(go.Scatter(x=monthly['주문월'], y=monthly['주문건수'], name='주문건수', line=dict(color='#E8853D', width=2.5), mode='lines+markers'), secondary_y=True)
    fig.update_layout(height=420, margin=dict(l=50, r=50, t=70, b=50),
                      title=dict(text='월별 매출 · 주문건수 추이', x=0.01, font=dict(size=15)),
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=11)))
    fig.update_yaxes(title_text="매출액", secondary_y=False)
    fig.update_yaxes(title_text="주문건수", secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)
    
    col_left, col_right = st.columns(2)
    with col_left:
        type_sales = filtered.groupby('주문자 구분')['판매합계금액'].sum().reset_index()
        type_sales.columns = ['구분', '매출']
        fig = px.pie(type_sales, values='매출', names='구분', hole=0.5, color_discrete_sequence=COLORS)
        fig.update_layout(height=420, title=dict(text='회원구분별 매출 비중', x=0.01, font=dict(size=15)),
                          margin=dict(l=20, r=20, t=70, b=20),
                          legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02, font=dict(size=11)))
        fig.update_traces(textinfo='percent+label', textfont_size=10)
        st.plotly_chart(fig, use_container_width=True)
    with col_right:
        region = filtered.groupby('지역')['판매합계금액'].sum().sort_values().reset_index()
        region.columns = ['지역', '매출']
        fig = px.bar(region, x='매출', y='지역', orientation='h', color_discrete_sequence=COLORS)
        fig.update_layout(height=420, title=dict(text='지역별 매출', x=0.01, font=dict(size=15)),
                          margin=dict(l=60, r=30, t=70, b=30), showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    daily = filtered.groupby('주문일자')['판매합계금액'].sum().reset_index()
    daily.columns = ['날짜', '매출']
    fig = px.area(daily, x='날짜', y='매출', color_discrete_sequence=['#3366CC'])
    fig.update_layout(height=350, title=dict(text='일별 매출 추이', x=0.01, font=dict(size=15)),
                      margin=dict(l=50, r=30, t=70, b=50), showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

# ============================================================
# Tab 2. 매출 분석
# ============================================================
with tab2:
    type_month = filtered.groupby(['주문월','주문자 구분'])['판매합계금액'].sum().reset_index()
    fig = px.bar(type_month, x='주문월', y='판매합계금액', color='주문자 구분', color_discrete_sequence=COLORS)
    fig.update_layout(height=420, barmode='stack',
                      title=dict(text='회원구분별 × 월별 매출 추이', x=0.01, font=dict(size=15)),
                      margin=dict(l=50, r=30, t=70, b=50),
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=10)))
    st.plotly_chart(fig, use_container_width=True)
    
    grade_sales = filtered.groupby('회원 등급')['판매합계금액'].sum().sort_values().reset_index()
    fig = px.bar(grade_sales, x='판매합계금액', y='회원 등급', orientation='h', color_discrete_sequence=COLORS)
    fig.update_layout(height=max(350, len(grade_sales)*28+100),
                      title=dict(text='회원등급별 매출', x=0.01, font=dict(size=15)),
                      margin=dict(l=130, r=30, t=70, b=30), showlegend=False)
    st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("##### 요일 · 시간대별 주문 매출 히트맵")
    dow_order = ['월','화','수','목','금','토','일']
    hm = filtered.groupby(['요일','주문시간'])['판매합계금액'].sum().reset_index()
    hm_pivot = hm.pivot_table(index='요일', columns='주문시간', values='판매합계금액', fill_value=0).reindex(dow_order)
    fig = go.Figure(data=go.Heatmap(z=hm_pivot.values, x=[f'{h}시' for h in hm_pivot.columns], y=hm_pivot.index,
                                      colorscale=[[0,'#F0F2F5'],[0.5,'#6B9BD2'],[1,'#1B2A4A']],
                                      hovertemplate='%{y} %{x}: %{z:,.0f}원<extra></extra>'))
    fig.update_layout(height=280, margin=dict(l=50, r=20, t=20, b=40))
    st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("##### 기관별 매출 현황")
    buyer_agg = filtered.groupby(['주문자 ID','주문자명','주문자 구분','회원 등급']).agg(
        매출=('판매합계금액','sum'), 주문건수=('주문 ID','nunique'), 최근주문일=('주문일자','max')).reset_index()
    buyer_agg['객단가'] = (buyer_agg['매출'] / buyer_agg['주문건수']).round(0)
    buyer_agg = buyer_agg.sort_values('매출', ascending=False)
    search = st.text_input("🔍 검색 (아이디, 주문자명)", key="sales_search")
    if search:
        buyer_agg = buyer_agg[buyer_agg.apply(lambda r: search.lower() in str(r).lower(), axis=1)]
    st.dataframe(buyer_agg.style.format({'매출':'{:,.0f}원','주문건수':'{:,.0f}건','객단가':'{:,.0f}원'}),
                 use_container_width=True, height=500)

# ============================================================
# Tab 3. 상품 분석
# ============================================================
with tab3:
    product_agg = filtered.groupby(['상품명','상품 코드']).agg(
        매출=('판매합계금액','sum'), 수량=('주문 수량','sum'), 주문건수=('주문 ID','nunique')
    ).reset_index().sort_values('매출', ascending=False)
    
    top20 = product_agg.head(20).copy()
    total_s = product_agg['매출'].sum()
    top20['누적비중'] = (top20['매출'].cumsum() / total_s * 100).round(1)
    
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=top20['상품명'].str[:18], y=top20['매출'], name='매출액', marker_color='#3366CC', opacity=0.8), secondary_y=False)
    fig.add_trace(go.Scatter(x=top20['상품명'].str[:18], y=top20['누적비중'], name='누적 비중',
                             line=dict(color='#E8853D', width=2.5), mode='lines+markers'), secondary_y=True)
    fig.update_layout(height=480, title=dict(text='상품별 매출 TOP 20 (파레토)', x=0.01, font=dict(size=15)),
                      margin=dict(l=50, r=50, t=70, b=130),
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=10)),
                      xaxis=dict(tickangle=45, tickfont=dict(size=8)))
    fig.update_yaxes(title_text="매출액", secondary_y=False)
    fig.update_yaxes(title_text="누적 비중 (%)", range=[0,105], secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("##### 전체 상품 매출 현황")
    search_p = st.text_input("🔍 상품명/코드 검색", key="product_search")
    dp = product_agg.copy()
    if search_p:
        dp = dp[dp.apply(lambda r: search_p.lower() in str(r).lower(), axis=1)]
    st.dataframe(dp.style.format({'매출':'{:,.0f}원','수량':'{:,.0f}','주문건수':'{:,.0f}건'}),
                 use_container_width=True, height=400)
    
    st.markdown("##### 회원구분별 × 상품 매출 크로스 (TOP 20)")
    t20n = product_agg.head(20)['상품명'].tolist()
    cp = filtered[filtered['상품명'].isin(t20n)].pivot_table(index='상품명', columns='주문자 구분', values='판매합계금액', aggfunc='sum', fill_value=0)
    cp['합계'] = cp.sum(axis=1)
    cp = cp.sort_values('합계', ascending=False)
    st.dataframe(cp.style.format('{:,.0f}원'), use_container_width=True, height=500)
    
    st.markdown("##### 월별 상품 매출 추이")
    t5 = product_agg.head(5)['상품명'].tolist()
    sel_p = st.multiselect("상품 선택", product_agg['상품명'].tolist(), default=t5, key="product_trend")
    if sel_p:
        td = filtered[filtered['상품명'].isin(sel_p)].groupby(['주문월','상품명'])['판매합계금액'].sum().reset_index()
        fig = px.line(td, x='주문월', y='판매합계금액', color='상품명', markers=True, color_discrete_sequence=COLORS)
        fig.update_layout(height=400, margin=dict(l=50, r=30, t=30, b=80),
                          legend=dict(orientation="h", yanchor="top", y=-0.18, x=0, font=dict(size=9)))
        for tr in fig.data:
            if len(tr.name) > 22: tr.name = tr.name[:22] + '...'
        st.plotly_chart(fig, use_container_width=True)

# ============================================================
# Tab 4. 회원 분석
# ============================================================
with tab4:
    mo = orders.groupby('주문자 ID').agg(첫주문일=('주문일','min'), 주문건수=('주문 ID','nunique'), 주문월수=('주문월','nunique')).reset_index()
    conv = members[members['아이디'].isin(orders['주문자 ID'].unique())]
    conv_rate = len(conv)/len(members)*100 if len(members)>0 else 0
    rep = mo[mo['주문건수']>=2]
    rep_rate = len(rep)/len(mo)*100 if len(mo)>0 else 0
    r3m = orders[orders['주문일']>=orders['주문일'].max()-pd.DateOffset(months=3)]
    act = r3m['주문자 ID'].nunique()
    
    cols = st.columns(5)
    for col, (l,v,u) in zip(cols, [
        ("총 회원수", fmt_number(len(members)), "처"), ("구매전환율", fmt_percent(conv_rate), ""),
        ("재구매율", fmt_percent(rep_rate), ""), ("활성회원", fmt_number(act), "처"),
        ("총 주문회원", fmt_number(len(mo)), "처"),
    ]):
        col.markdown(kpi_card(l,v,u), unsafe_allow_html=True)
    
    cl, cr = st.columns(2)
    with cl:
        jm = members.groupby(['가입월','회원타입']).size().reset_index(name='가입자수')
        fig = px.bar(jm, x='가입월', y='가입자수', color='회원타입', color_discrete_sequence=COLORS)
        fig.update_layout(height=420, barmode='stack',
                          title=dict(text='월별 신규가입자 추이', x=0.01, font=dict(size=15)),
                          margin=dict(l=50, r=20, t=70, b=60),
                          legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=10)))
        st.plotly_chart(fig, use_container_width=True)
    with cr:
        gd = members['회원등급'].value_counts().reset_index()
        gd.columns = ['등급','수']
        fig = px.bar(gd, x='수', y='등급', orientation='h', color_discrete_sequence=COLORS)
        fig.update_layout(height=420, title=dict(text='회원등급별 분포', x=0.01, font=dict(size=15)),
                          margin=dict(l=130, r=20, t=70, b=30), showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    cl, cr = st.columns(2)
    with cl:
        mg = members.merge(mo, left_on='아이디', right_on='주문자 ID', how='inner')
        mg['소요일'] = (mg['첫주문일']-mg['가입일']).dt.days
        mg = mg[mg['소요일']>=0]
        bins=[0,1,8,15,31,61,91,9999]; lb=['당일','1~7일','8~14일','15~30일','31~60일','61~90일','90일+']
        mg['구간'] = pd.cut(mg['소요일'], bins=bins, labels=lb, right=False)
        dh = mg['구간'].value_counts().reindex(lb).fillna(0).reset_index()
        dh.columns=['구간','회원수']
        fig = px.bar(dh, x='구간', y='회원수', color_discrete_sequence=['#3366CC'])
        fig.update_layout(height=400, title=dict(text='첫 주문까지 소요일', x=0.01, font=dict(size=15)),
                          margin=dict(l=50, r=20, t=70, b=50), showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    with cr:
        bo=[1,2,4,6,11,21,9999]; lo=['1회','2~3회','4~5회','6~10회','11~20회','20회+']
        mo['구간'] = pd.cut(mo['주문건수'], bins=bo, labels=lo, right=False)
        od = mo['구간'].value_counts().reindex(lo).fillna(0).reset_index()
        od.columns=['구간','회원수']
        fig = px.bar(od, x='구간', y='회원수', color_discrete_sequence=COLORS)
        fig.update_layout(height=400, title=dict(text='주문횟수별 회원 분포', x=0.01, font=dict(size=15)),
                          margin=dict(l=50, r=20, t=70, b=50), showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("##### 코호트 리텐션 히트맵")
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
        rdf=pd.DataFrame(rd); mc=[f'{i}개월' for i in range(mx)]; zv=rdf[mc].values
        fig=go.Figure(data=go.Heatmap(z=zv, x=mc,
            y=[f"{r['코호트']} ({r['크기']}처)" for _,r in rdf.iterrows()],
            colorscale=[[0,'#F0F2F5'],[0.3,'#A8D5A2'],[1,'#27AE60']],
            text=[[f'{v:.1f}%' if v>0 else '-' for v in row] for row in zv],
            texttemplate='%{text}', textfont=dict(size=9),
            hovertemplate='%{y}<br>%{x}: %{z:.1f}%<extra></extra>'))
        fig.update_layout(height=max(350,len(rd)*30+120), margin=dict(l=160,r=20,t=20,b=40),
                          yaxis=dict(tickfont=dict(size=9), autorange="reversed"))
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
        if n not in ra:
            ra[n]={'추천인':n,'유형':clm.get(n,'케어포'),'추천인코드':r.get('추천인코드',''),'피추천인수':0,'biz':[]}
        b=str(r.get('피추천인 사업자 번호','')).strip()
        if b and b not in ['-','nan']:
            ra[n]['피추천인수']+=1; ra[n]['biz'].append(b)
    
    b2u=members.set_index('사업자번호')['아이디'].to_dict()
    bs=filtered.groupby('주문자 ID')['판매합계금액'].sum().to_dict()
    for n in ra:
        ra[n]['피추천인매출']=sum(bs.get(b2u.get(b,''),0) for b in ra[n]['biz'])
    
    rdf=pd.DataFrame(ra.values())[['추천인','유형','추천인코드','피추천인수','피추천인매출']]
    
    cols=st.columns(3)
    for col,(l,v,u) in zip(cols,[
        ("총 추천인 수",fmt_number(len(rdf)),"회원"),
        ("총 피추천인 수",fmt_number(rdf['피추천인수'].sum()),"회원"),
        ("추천인당 평균",f"{rdf['피추천인수'].mean():.1f}" if len(rdf)>0 else "0","회원"),
    ]):
        col.markdown(kpi_card(l,v,u), unsafe_allow_html=True)
    
    tc={'영업팀':'#3366CC','대리점':'#E8853D','케어포':'#27AE60'}
    cl,cr=st.columns(2)
    with cl:
        tr=rdf.groupby('유형')['피추천인수'].sum().reset_index()
        fig=px.bar(tr, x='유형', y='피추천인수', color='유형', color_discrete_map=tc)
        fig.update_layout(height=400, showlegend=False,
                          title=dict(text='유형별 피추천인 수', x=0.01, font=dict(size=15)),
                          margin=dict(l=50,r=20,t=70,b=30))
        st.plotly_chart(fig, use_container_width=True)
    with cr:
        ts=rdf.groupby('유형')['피추천인매출'].sum().reset_index()
        fig=px.pie(ts, values='피추천인매출', names='유형', hole=0.5, color='유형', color_discrete_map=tc)
        fig.update_layout(height=400, title=dict(text='유형별 피추천인 매출', x=0.01, font=dict(size=15)),
                          margin=dict(l=20,r=20,t=70,b=20))
        st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("##### 추천인별 현황")
    rtf=st.selectbox("유형 필터",["전체","영업팀","대리점","케어포"],key="ref_type")
    dr=rdf.copy()
    if rtf!="전체": dr=dr[dr['유형']==rtf]
    dr=dr.sort_values('피추천인매출',ascending=False)
    sr=st.text_input("🔍 추천인 검색",key="ref_search")
    if sr: dr=dr[dr.apply(lambda r:sr.lower() in str(r).lower(),axis=1)]
    st.dataframe(dr.style.format({'피추천인수':'{:,.0f}','피추천인매출':'{:,.0f}원'}),
                 use_container_width=True, height=500)

# ============================================================
# Tab 6. 케어포 멤버십
# ============================================================
with tab6:
    cfg=['케어포-시설','케어포-공생','케어포-주야간','케어포-방문','케어포-일반','케어포-종사자','케어포-보호자']
    co=filtered[filtered['회원 등급'].isin(cfg)]
    cmb=members[members['회원타입']=='케어포']
    
    cgf=st.selectbox("케어포 등급",["전체"]+cfg,key="cf_grade")
    if cgf!="전체": co=co[co['회원 등급']==cgf]
    
    cbo=co.groupby('주문자 ID')['주문 ID'].nunique()
    crp=(cbo>=2).sum(); crr=crp/len(cbo)*100 if len(cbo)>0 else 0
    
    cols=st.columns(3)
    for col,(l,v,u) in zip(cols,[
        ("케어포 총 회원",fmt_number(len(cmb)),"처"),
        ("케어포 주문회원",fmt_number(co['주문자 ID'].nunique()),"처"),
        ("케어포 재구매율",fmt_percent(crr),""),
    ]):
        col.markdown(kpi_card(l,v,u), unsafe_allow_html=True)
    
    cl,cr=st.columns(2)
    with cl:
        cga=co.groupby('회원 등급').agg(매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique')).reset_index()
        cga['등급']=cga['회원 등급'].str.replace('케어포-','')
        fig=make_subplots(specs=[[{"secondary_y":True}]])
        fig.add_trace(go.Bar(x=cga['등급'],y=cga['매출'],name='매출액',marker_color='#3366CC',opacity=0.8),secondary_y=False)
        fig.add_trace(go.Scatter(x=cga['등급'],y=cga['주문건수'],name='주문건수',line=dict(color='#E8853D',width=2.5),mode='lines+markers'),secondary_y=True)
        fig.update_layout(height=420, title=dict(text='등급별 매출·주문', x=0.01, font=dict(size=15)),
                          margin=dict(l=50,r=50,t=70,b=30),
                          legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=10)))
        st.plotly_chart(fig, use_container_width=True)
    with cr:
        cpd=co[co['상품명'].str.contains(r'\[케어포',na=False)]
        cpm=cpd.groupby('주문월')['판매합계금액'].sum().reset_index()
        fig=px.area(cpm,x='주문월',y='판매합계금액',color_discrete_sequence=['#27AE60'])
        fig.update_layout(height=420, title=dict(text='전용 상품 매출 추이', x=0.01, font=dict(size=15)),
                          margin=dict(l=50,r=30,t=70,b=50), showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    cj=cmb.groupby(['가입월','회원등급']).size().reset_index(name='가입자수')
    ccl={'케어포-시설':'#3366CC','케어포-공생':'#E8853D','케어포-주야간':'#27AE60',
         '케어포-방문':'#9B59B6','케어포-일반':'#F39C12','케어포-종사자':'#E74C3C','케어포-보호자':'#1ABC9C'}
    fig=px.bar(cj,x='가입월',y='가입자수',color='회원등급',color_discrete_map=ccl)
    fig.update_layout(height=420, barmode='stack',
                      title=dict(text='등급별 신규가입 추이', x=0.01, font=dict(size=15)),
                      margin=dict(l=50,r=20,t=70,b=80),
                      legend=dict(orientation="h",yanchor="top",y=-0.15,x=0,font=dict(size=9)))
    st.plotly_chart(fig, use_container_width=True)

# ============================================================
# 푸터
# ============================================================
st.markdown("---")
st.markdown(f"<p style='text-align:center;color:#94a3b8;font-size:0.8rem;'>© 대상웰라이프 B2B몰 대시보드 · 데이터 기준: {pd.Timestamp.now().strftime('%Y-%m-%d')}</p>", unsafe_allow_html=True)
