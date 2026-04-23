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
    sign = '' if n >= 0 else '-'; a = abs(n)
    if a >= 1e8: return f"{sign}{a/1e8:.1f}억원"
    if a >= 1e4: return f"{sign}{a/1e4:,.0f}만원"
    return f"{sign}{a:,.0f}원"
def fmt_krw_short(n):
    if pd.isna(n) or n == 0: return "0"
    sign = '' if n >= 0 else '-'; a = abs(n)
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
def ym_series_kr(series): return series.apply(to_ym_kr)
def to_date_kr(date_str):
    if not date_str or pd.isna(date_str): return ''
    parts = str(date_str).split('-')
    if len(parts) == 3: return f"{parts[0]}년 {int(parts[1])}월 {int(parts[2])}일"
    return str(date_str)
def krw_tickvals(series, n=5):
    mn, mx = series.min(), series.max()
    if mx == 0: return [0], ['0']
    vals = np.linspace(0, mx * 1.05, n).tolist()
    return vals, [fmt_krw_short(v) for v in vals]
def make_donut(df, name_col, value_col, colors=None, value_label='매출', unit='원'):
    total = df[value_col].sum()
    fig = go.Figure()
    fig.add_trace(go.Pie(labels=df[name_col], values=df[value_col], hole=0.55,
        marker=dict(colors=(colors or COLORS)[:len(df)]),
        textinfo='label+percent', textposition='inside', insidetextorientation='horizontal', textfont=dict(size=11),
        hovertemplate='%{label}<br>' + value_label + ': %{customdata}<br>비중: %{percent}<extra></extra>',
        customdata=[f"{v:,.0f}{unit}" for v in df[value_col]]))
    total_fmt = f"{total:,.0f}{unit}" if unit != '원' else fmt_krw(total)
    fig.add_annotation(text=f"<b>합계</b><br>{total_fmt}", x=0.5, y=0.5, font=dict(size=15, color='#1e293b'), showarrow=False, xref='paper', yref='paper')
    fig.update_layout(height=520, margin=dict(l=20, r=20, t=30, b=140),
        legend=dict(orientation="h", yanchor="top", y=-0.02, xanchor="center", x=0.5, font=dict(size=11)), showlegend=True)
    return fig

# key 헬퍼: kp가 있으면 고유 key, 없으면 None
def _k(kp, suffix):
    return f"{kp}_{suffix}" if kp else None

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
    orders['구매확정일'] = pd.to_datetime(orders['구매확정일'], errors='coerce')
    def calc_settlement_month(dt):
        if pd.isna(dt): return None
        if dt.day >= 26:
            # 다음달로 이월
            if dt.month == 12: return f"{dt.year+1}-01"
            else: return f"{dt.year}-{str(dt.month+1).zfill(2)}"
        else:
            return f"{dt.year}-{str(dt.month).zfill(2)}"
    orders['정산월'] = orders['구매확정일'].apply(calc_settlement_month)
    members['가입일'] = pd.to_datetime(members['가입일'], errors='coerce')
    members['가입월'] = members['가입일'].dt.to_period('M').astype(str)
    members['사업자번호'] = members['사업자번호'].astype(str).str.replace('-','').str.strip()
    referrals_df['피추천인 사업자 번호'] = referrals_df['피추천인 사업자 번호'].astype(str).str.replace('-','').str.strip()
    return orders, members, referrals_df

@st.cache_data
def process_bw(bw_raw):
    bw = bw_raw.copy()
    bw.columns = [c.replace('\n','').strip() for c in bw.columns]
    def parse_ym(val):
        if pd.isna(val): return None
        s = str(val).strip()
        if '.' in s:
            parts = s.split('.')
            return f"{parts[0]}-{parts[1].zfill(2)}"
        return s
    bw['연월'] = bw['달력 연도/월'].apply(parse_ym)
    bw['연도'] = bw['연월'].str[:4]
    bw['월'] = bw['연월'].str[5:7].astype(int, errors='ignore')
    def customer_label(name):
        s = str(name).strip(); parts = s.split(',')
        if len(parts) == 1: return '일반'
        ref_map = {'영':'영업','대':'대리점','케':'케어포'}
        mem_map = {'의':'의료기','장':'장기요양','병':'병원','약':'약국','크':'염증성장질환','종':'종사자'}
        if len(parts) == 2: return ref_map.get(parts[1].strip(), parts[1].strip())
        if len(parts) == 3:
            p1, p2 = parts[1].strip(), parts[2].strip()
            if p2 == '종': return f"{mem_map.get(p1, ref_map.get(p1, p1))}-종사자"
            return f"{ref_map.get(p1, p1)}-{mem_map.get(p2, p2)}"
        return s
    bw['채널'] = bw['고객명'].apply(customer_label)
    bw['기타판관비'] = bw['IV.판매비 및 관리비'] - bw['IV.6.광고선전비'] - bw['IV.7.운반비'] - bw['IV.8.판매수수료'] - bw['IV.9.판촉비']
    bw['매출총이익률'] = np.where(bw['I.매출액(FI기준)'] != 0, bw['III.매출총이익'] / bw['I.매출액(FI기준)'] * 100, 0)
    bw['영업이익률'] = np.where(bw['I.매출액(FI기준)'] != 0, bw['V.영업이익I'] / bw['I.매출액(FI기준)'] * 100, 0)
    return bw

# ============================================================
# 일차의료 시범기관
# ============================================================
PILOT_SHEET_ID = '1ZV_Rxi6FGuzazcm0FFKcXQ7Xa1KwD3DRjo1KY8W6EuY'
PILOT_SHEETS = {'만성질환관리':'만성질환관리','방문진료':'일차의료_방문진료','한의방문진료':'일차의료_한의_방문진료'}
SIDO_NORM = {
    '서울특별시':'서울','서울시':'서울','서울':'서울','부산광역시':'부산','부산시':'부산','부산':'부산',
    '대구광역시':'대구','대구시':'대구','대구':'대구','인천광역시':'인천','인천시':'인천','인천':'인천',
    '광주광역시':'광주','광주시':'광주','광주':'광주','대전광역시':'대전','대전시':'대전','대전':'대전',
    '울산광역시':'울산','울산시':'울산','울산':'울산','세종특별자치시':'세종','세종시':'세종','세종':'세종',
    '경기도':'경기','경기':'경기','강원특별자치도':'강원','강원도':'강원','강원':'강원',
    '충청북도':'충북','충북':'충북','충청남도':'충남','충남':'충남',
    '전북특별자치도':'전북','전라북도':'전북','전북':'전북','전라남도':'전남','전남':'전남',
    '경상북도':'경북','경북':'경북','경상남도':'경남','경남':'경남',
    '제주특별자치도':'제주','제주도':'제주','제주':'제주',
}
def normalize_sido(addr):
    if pd.isna(addr) or not addr: return ''
    first = str(addr).strip().split()[0] if str(addr).strip() else ''
    return SIDO_NORM.get(first, first)
def extract_sigungu(addr):
    if pd.isna(addr) or not addr: return ''
    parts = str(addr).strip().split()
    return parts[1] if len(parts) > 1 else ''
def extract_road(addr):
    if pd.isna(addr) or not addr: return ''
    parts = str(addr).strip().split(); road_idx = 2
    for i in range(2, min(len(parts), 5)):
        word = parts[i].rstrip(',')
        if word.endswith(('구','군')): road_idx = i + 1; continue
        if word.endswith(('로','길')): road_idx = i; break
    if road_idx < len(parts):
        road = parts[road_idx].rstrip(',')
        if road_idx + 1 < len(parts): return road + ' ' + parts[road_idx+1].rstrip(',')
        return road
    return ''
def normalize_name(name):
    if pd.isna(name) or not name: return ''
    import re; s = str(name).strip()
    s = re.sub(r'[\s\(\)·\-\.\,\']', '', s)
    for prefix in ['의료법인','사회복지법인','재단법인','학교법인']: s = s.replace(prefix, '')
    return s
def name_similarity(a, b):
    if not a or not b: return 0
    import re
    if a == b: return 100
    suffixes = r'(한의원|의원|병원|클리닉|약국|요양원|의료원|보건소|한방병원|치과)$'
    a_s = re.sub(suffixes, '', a); b_s = re.sub(suffixes, '', b)
    if a_s and b_s and a_s == b_s: return 98
    dept = r'(내과|치과|외과|피부과|안과|이비인후과|정형외과|산부인과|비뇨기과|신경과|정신과|재활의학과|소아과|한의원)'
    if re.findall(dept,a) and re.findall(dept,b) and re.findall(dept,a) != re.findall(dept,b): return 0
    if a in b or b in a: return int(min(len(a),len(b))/max(len(a),len(b))*100)
    common = len(set(a)&set(b)); total = max(len(set(a)),len(set(b)))
    if total == 0: return 0
    base = int(common/total*100); prefix_len = 0
    for ca, cb in zip(a, b):
        if ca == cb: prefix_len += 1
        else: break
    if prefix_len >= 2: base = min(100, base + prefix_len * 5)
    return base

@st.cache_data(ttl=3600, show_spinner="🏥 일차의료 시범기관 데이터를 불러오는 중...")
def load_pilot_clinics():
    from urllib.parse import quote
    all_frames = []
    for label, sheet_name in PILOT_SHEETS.items():
        try:
            url = f"https://docs.google.com/spreadsheets/d/{PILOT_SHEET_ID}/gviz/tq?tqx=out:csv&sheet={quote(sheet_name)}"
            df = pd.read_csv(url); df.columns = [c.strip() for c in df.columns]
            rename_map = {'NO':'no','병원/약국명':'기관명','병원/약국구분':'기관구분','전화번호':'전화번호','우편번호':'우편번호','소재지주소':'주소','홈페이지':'홈페이지'}
            df = df.rename(columns=rename_map); df['사업유형'] = label; all_frames.append(df)
        except Exception as e:
            st.error(f"시트 '{sheet_name}' 로드 실패: {str(e)}"); continue
    if not all_frames: return pd.DataFrame()
    result = pd.concat(all_frames, ignore_index=True)
    result['전화번호_norm'] = result['전화번호'].fillna('').astype(str).str.replace(r'[^0-9]','',regex=True)
    result['기관명_norm'] = result['기관명'].apply(normalize_name)
    result['시도'] = result['주소'].apply(normalize_sido)
    result['시군구'] = result['주소'].apply(extract_sigungu)
    result['도로명'] = result['주소'].apply(extract_road)
    return result

def match_pilot_clinics(pilot_df, members_df, orders_df):
    if pilot_df.empty or members_df.empty: return pd.DataFrame()
    MANUAL_MATCH = {
        ('건강드림내과의원','경기','평택시'):'happydream',('구로연세의원','서울','구로구'):'kuroyonsei',
        ('상인내과의원','대구','달서구'):'sangin',('서울배내과의원','서울','강남구'):'flowerbae',
        ('성모가정의학과의원','인천','남동구'):'fmkjm4',('아세아연합의원','대구','서구'):'hyyk1213',
        ('웰봄내과의원','경기','오산시'):'wbch01739',('이시아연합속내과의원','대구','동구'):'stereon',
        ('차만진가정의학과의원','광주','북구'):'sssky91',('참사랑내과의원','경남','하동군'):'jhs7575',
        ('첨단가족연합의원','광주','북구'):'tontokim',('하늘내과의원','전북','전주시'):'skymed0813',
        ('한양류마유내과의원','경기','화성시'):'hyrmu',
    }
    mem = members_df.copy()
    mem = mem[(mem['회원타입']=='병원')&(mem['회원등급']=='병원')]
    mem['전화번호_norm'] = mem['휴대폰'].fillna('').astype(str).str.replace(r'[^0-9]','',regex=True)
    mem['상호명_norm'] = mem['상호명'].apply(normalize_name)
    mem['시도'] = mem['주소'].apply(normalize_sido)
    mem['시군구'] = mem['주소'].apply(extract_sigungu)
    mem['도로명'] = mem['주소'].apply(extract_road)
    order_agg = orders_df.groupby('주문자 ID').agg(총매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique'),최근주문일=('주문일자','max')).reset_index()
    results = []; matched_idx = set(); mem_by_id = mem.set_index('아이디')
    for idx, row in pilot_df.iterrows():
        bid = MANUAL_MATCH.get((row['기관명'],row['시도'],row['시군구']))
        if bid and bid in mem_by_id.index:
            m = mem_by_id.loc[bid]
            results.append({'기관명':row['기관명'],'기관구분':row.get('기관구분',''),'사업유형':row['사업유형'],'주소_공공':row['주소'],'전화번호_공공':row['전화번호'],'상호명_B2B':m['상호명'],'아이디':bid,'주소_B2B':m['주소'],'회원타입':m.get('회원타입',''),'회원등급':m.get('회원등급',''),'매칭방법':'수동매칭','매칭등급':'확정','유사도':100})
            matched_idx.add(idx)
    phone_map = {}
    for _, m in mem[mem['전화번호_norm'].str.len()>=9].iterrows(): phone_map.setdefault(m['전화번호_norm'], m)
    for idx, row in pilot_df.iterrows():
        phone = row['전화번호_norm']
        if phone and len(phone)>=9 and phone in phone_map:
            m = phone_map[phone]
            results.append({'기관명':row['기관명'],'기관구분':row.get('기관구분',''),'사업유형':row['사업유형'],'주소_공공':row['주소'],'전화번호_공공':row['전화번호'],'상호명_B2B':m['상호명'],'아이디':m['아이디'],'주소_B2B':m['주소'],'회원타입':m.get('회원타입',''),'회원등급':m.get('회원등급',''),'매칭방법':'전화번호','매칭등급':'확정','유사도':100})
            matched_idx.add(idx)
    mem_region_name = {}
    for _, m in mem.iterrows():
        key = (m['시도'],m['시군구'],m['상호명_norm'])
        if key not in mem_region_name: mem_region_name[key] = m
    for idx, row in pilot_df.iterrows():
        if idx in matched_idx: continue
        key = (row['시도'],row['시군구'],row['기관명_norm'])
        if key in mem_region_name:
            m = mem_region_name[key]
            results.append({'기관명':row['기관명'],'기관구분':row.get('기관구분',''),'사업유형':row['사업유형'],'주소_공공':row['주소'],'전화번호_공공':row['전화번호'],'상호명_B2B':m['상호명'],'아이디':m['아이디'],'주소_B2B':m['주소'],'회원타입':m.get('회원타입',''),'회원등급':m.get('회원등급',''),'매칭방법':'상호명+지역','매칭등급':'확정','유사도':100})
            matched_idx.add(idx)
    mem_by_region = {}
    for _, m in mem.iterrows():
        rkey = (m['시도'],m['시군구'],m['도로명']); mem_by_region.setdefault(rkey,[]).append(m)
    for idx, row in pilot_df.iterrows():
        if idx in matched_idx: continue
        cname = row['기관명_norm']; rkey = (row['시도'],row['시군구'],extract_road(row['주소']))
        if not cname or rkey not in mem_by_region: continue
        best_score = 0; best_m = None
        for m in mem_by_region[rkey]:
            score = name_similarity(cname, m['상호명_norm'])
            if score > best_score: best_score = score; best_m = m
        if best_score >= 85 and best_m is not None:
            results.append({'기관명':row['기관명'],'기관구분':row.get('기관구분',''),'사업유형':row['사업유형'],'주소_공공':row['주소'],'전화번호_공공':row['전화번호'],'상호명_B2B':best_m['상호명'],'아이디':best_m['아이디'],'주소_B2B':best_m['주소'],'회원타입':best_m.get('회원타입',''),'회원등급':best_m.get('회원등급',''),'매칭방법':f'유사도매칭({best_score}%)','매칭등급':'확정','유사도':best_score})
            matched_idx.add(idx)
    if not results: return pd.DataFrame()
    result_df = pd.DataFrame(results)
    result_df = result_df.merge(order_agg, left_on='아이디', right_on='주문자 ID', how='left').drop(columns=['주문자 ID'],errors='ignore')
    result_df['총매출'] = result_df['총매출'].fillna(0)
    result_df['주문건수'] = result_df['주문건수'].fillna(0).astype(int)
    result_df['최근주문일'] = result_df['최근주문일'].fillna('-')
    return result_df.sort_values('유사도', ascending=False)

# ============================================================
# 데이터 로드
# ============================================================
GDRIVE_FILE_ID = '1Op9Y2FFb_aLQJKAcLyKj9HJQbK6YYnmf'
def download_from_gdrive(file_id):
    import gdown, tempfile, os
    tmp = tempfile.mktemp(suffix='.xlsx')
    gdown.download(f'https://drive.google.com/uc?id={file_id}', tmp, quiet=True)
    with open(tmp, 'rb') as f: content = f.read()
    os.remove(tmp)
    return io.BytesIO(content)

@st.cache_data(ttl=3600, show_spinner="📥 구글 드라이브에서 데이터를 불러오는 중...", max_entries=1)
def load_from_gdrive():
    fb = download_from_gdrive(GDRIVE_FILE_ID)
    o = pd.read_excel(fb, sheet_name='주문내역', header=1, engine='openpyxl'); fb.seek(0)
    m = pd.read_excel(fb, sheet_name='회원정보', header=1, engine='openpyxl'); fb.seek(0)
    r = pd.read_excel(fb, sheet_name='추천인', header=1, engine='openpyxl'); fb.seek(0)
    try:
        dealer_raw = pd.read_excel(fb, sheet_name='대리점 피추천인 주문내역', header=1, engine='openpyxl')
    except Exception:
        dealer_raw = pd.DataFrame()
    fb.seek(0)
    try:
        bw_raw = pd.read_excel(fb, sheet_name='BW', header=0, engine='openpyxl', dtype={'달력 연도/월':str})
        bw = process_bw(bw_raw)
    except Exception: bw = pd.DataFrame()
    orders, members, referrals_df = process_data(o, m, r)
    return orders, members, referrals_df, bw, dealer_raw

try:
    orders, members, referrals_df, bw_data, dealer_raw = load_from_gdrive()
    sidebar_msg = f"✅ 데이터 로드 완료\n- 주문: {len(orders):,}건\n- 회원: {len(members):,}건\n- 추천인: {len(referrals_df):,}건"
    if len(bw_data) > 0: sidebar_msg += f"\n- BW손익: {len(bw_data):,}건"
    st.sidebar.success(sidebar_msg)
except Exception as e:
    st.error(f"❌ 데이터 로드 실패: {str(e)}"); st.stop()
if st.sidebar.button("🔄 데이터 새로고침"):
    st.cache_data.clear(); st.rerun()
st.sidebar.markdown("---")
member_lookup = members.set_index('아이디')[['상호명','회원타입','회원등급']].to_dict('index')

st.sidebar.markdown("## 🔍 필터")
years = sorted(orders['주문일'].dt.year.dropna().unique().astype(int))
selected_years = st.sidebar.multiselect("연도", [str(y) for y in years], default=[], placeholder="전체")
if selected_years:
    all_months = set()
    for y in selected_years:
        ms = orders[orders['주문일'].dt.year==int(y)]['주문일'].dt.month.dropna().unique().astype(int)
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
bw_filtered = bw_data.copy()
if len(bw_filtered) > 0:
    if selected_years: bw_filtered = bw_filtered[bw_filtered['연도'].isin(selected_years)]
    if selected_months: bw_filtered = bw_filtered[bw_filtered['월'].isin([int(m.replace('월','')) for m in selected_months])]

import base64, os
logo_path = os.path.join(os.path.dirname(__file__), 'logo.png')
if os.path.exists(logo_path):
    with open(logo_path, 'rb') as f: logo_b64 = base64.b64encode(f.read()).decode()
    st.markdown(f'<div class="main-header" style="display:flex;align-items:center;justify-content:space-between;"><div><h1>📊 대상웰라이프 B2B몰 대시보드</h1><p>Sales & Operations Analytics</p></div><img src="data:image/png;base64,{logo_b64}" style="height:50px;object-fit:contain;"></div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="main-header"><h1>📊 대상웰라이프 B2B몰 대시보드</h1><p>Sales & Operations Analytics</p></div>', unsafe_allow_html=True)

# ============================================================
# render 함수들 — kp="" 이면 일반탭, kp="custom_C01_0" 이면 커스터마이징탭
# ============================================================

def render_monthly_sales_trend(kp=""):
    st.markdown("#### 월별 매출 · 주문건수 추이")
    monthly = filtered.groupby('주문월').agg(매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique')).reset_index()
    monthly['주문월_kr'] = ym_series_kr(monthly['주문월'])
    fig = make_subplots(specs=[[{"secondary_y":True}]])
    fig.add_trace(go.Bar(x=monthly['주문월_kr'],y=monthly['매출'],name='매출액',marker_color='#3366CC',opacity=0.8,text=[fmt_krw_short(v) for v in monthly['매출']],textposition='outside',textfont=dict(size=11),hovertemplate='%{x}<br>매출: %{customdata}<extra></extra>',customdata=[f"{v:,.0f}원" for v in monthly['매출']]),secondary_y=False)
    fig.add_trace(go.Scatter(x=monthly['주문월_kr'],y=monthly['주문건수'],name='주문건수',line=dict(color='#E8853D',width=3),mode='lines+markers',marker=dict(size=8),hovertemplate='%{x}<br>주문: %{y:,}건<extra></extra>'),secondary_y=True)
    tvals,ttexts = krw_tickvals(monthly['매출'])
    fig.update_layout(height=480,margin=dict(l=80,r=60,t=50,b=70),legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=12)),xaxis=dict(tickfont=dict(size=12)))
    fig.update_yaxes(title_text="매출액",tickvals=tvals,ticktext=ttexts,tickfont=dict(size=11),secondary_y=False)
    fig.update_yaxes(title_text="주문건수",tickfont=dict(size=11),secondary_y=True)
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"monthly_sales"))

def render_member_type_sales_donut(kp=""):
    st.markdown("#### 회원구분별 매출 비중")
    ts_df = filtered.groupby('주문자 구분')['판매합계금액'].sum().reset_index(); ts_df.columns=['구분','매출']; ts_df=ts_df.sort_values('매출',ascending=False)
    fig = make_donut(ts_df,'구분','매출'); fig.update_layout(height=520)
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"type_donut"))

def render_region_sales_bar(kp=""):
    st.markdown("#### 지역별 매출")
    rg = filtered.groupby('지역')['판매합계금액'].sum().sort_values().reset_index(); rg.columns=['지역','매출']
    fig = px.bar(rg,x='매출',y='지역',orientation='h',color_discrete_sequence=COLORS)
    fig.update_traces(text=[fmt_krw_short(v) for v in rg['매출']],textposition='outside',textfont=dict(size=11),hovertemplate='%{y}: %{customdata}<extra></extra>',customdata=[f"{v:,.0f}원" for v in rg['매출']])
    fig.update_layout(height=520,margin=dict(l=70,r=100,t=30,b=40),showlegend=False)
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"region_bar"))

def render_daily_sales_trend(kp=""):
    st.markdown("#### 일별 매출 추이")
    daily = filtered.groupby('주문일자')['판매합계금액'].sum().reset_index(); daily.columns=['날짜','매출']
    fig = px.area(daily,x='날짜',y='매출',color_discrete_sequence=['#3366CC'])
    fig.update_traces(customdata=list(zip(daily['날짜'].apply(to_date_kr),[f"{v:,.0f}원" for v in daily['매출']])),hovertemplate='%{customdata[0]}<br>매출: %{customdata[1]}<extra></extra>')
    tvals2,ttexts2 = krw_tickvals(daily['매출'])
    fig.update_layout(height=400,margin=dict(l=80,r=30,t=30,b=60),showlegend=False,xaxis=dict(title='날짜',tickfont=dict(size=11),tickformat='%Y년 %m월'),yaxis=dict(title='매출액',tickvals=tvals2,ticktext=ttexts2,tickfont=dict(size=11)))
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"daily_sales"))

def render_type_monthly_sales(kp=""):
    st.markdown("#### 회원구분별 × 월별 매출 추이")
    tm_df = filtered.groupby(['주문월','주문자 구분'])['판매합계금액'].sum().reset_index(); tm_df['주문월_kr']=ym_series_kr(tm_df['주문월'])
    fig = px.bar(tm_df,x='주문월_kr',y='판매합계금액',color='주문자 구분',color_discrete_sequence=COLORS)
    for tr in fig.data: tr.customdata=[f"{v:,.0f}원" for v in tr.y]; tr.hovertemplate='%{x}<br>'+tr.name+': %{customdata}<extra></extra>'
    fig.update_layout(height=480,barmode='stack',margin=dict(l=70,r=30,t=50,b=70),legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=11)))
    total_by_month = tm_df.groupby('주문월_kr')['판매합계금액'].sum().reset_index()
    fig.add_trace(go.Scatter(x=total_by_month['주문월_kr'], y=total_by_month['판매합계금액'], mode='text', text=[fmt_krw_short(v) for v in total_by_month['판매합계금액']], textposition='top center', textfont=dict(size=11, color='#1e293b'), showlegend=False, hoverinfo='skip'))
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"type_monthly"))

def render_grade_sales_bar(kp=""):
    st.markdown("#### 회원등급별 매출")
    gs = filtered.groupby('회원 등급').agg(매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique'),주문회원수=('주문자 ID','nunique')).reset_index().sort_values('매출')
    fig = px.bar(gs,x='매출',y='회원 등급',orientation='h',color_discrete_sequence=COLORS)
    fig.update_traces(text=[fmt_krw_short(v) for v in gs['매출']],textposition='outside',textfont=dict(size=11),hovertemplate='%{y}<br>매출: %{customdata[0]}<br>주문: %{customdata[1]:,}건<br>회원: %{customdata[2]:,}처<extra></extra>',customdata=list(zip([f"{v:,.0f}원" for v in gs['매출']],gs['주문건수'],gs['주문회원수'])))
    fig.update_layout(height=max(420,len(gs)*35+140),margin=dict(l=140,r=100,t=30,b=40),showlegend=False)
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"grade_sales"))

def render_heatmap_dow_hour(kp=""):
    st.markdown("#### 요일 · 시간대별 주문 매출 히트맵")
    dow_order = ['월','화','수','목','금','토','일']
    hm = filtered.groupby(['요일','주문시간'])['판매합계금액'].sum().reset_index()
    hmp = hm.pivot_table(index='요일',columns='주문시간',values='판매합계금액',fill_value=0).reindex(dow_order)
    fig = go.Figure(data=go.Heatmap(z=hmp.values,x=[f'{h}시' for h in hmp.columns],y=hmp.index,colorscale=[[0,'#F0F2F5'],[0.5,'#6B9BD2'],[1,'#1B2A4A']],text=[[fmt_krw_short(v) for v in row] for row in hmp.values],texttemplate='%{text}',textfont=dict(size=12),hovertemplate='%{y} %{x}<br>매출: %{customdata}<extra></extra>',customdata=[[f"{v:,.0f}원" for v in row] for row in hmp.values]))
    fig.update_layout(height=320,margin=dict(l=50,r=20,t=20,b=40),xaxis=dict(tickfont=dict(size=11)),yaxis=dict(tickfont=dict(size=12),autorange='reversed'))
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"heatmap"))

def render_org_sales_table(kp=""):
    st.markdown("#### 기관별 매출 현황")
    ba = filtered.groupby(['주문자 ID','주문자명','주문자 구분','회원 등급']).agg(매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique'),최근주문일=('주문일자','max')).reset_index()
    ba['객단가']=(ba['매출']/ba['주문건수']).round(0); ba['상호명']=ba['주문자 ID'].map(lambda x:member_lookup.get(x,{}).get('상호명',''))
    ba=ba[['주문자 ID','주문자명','상호명','주문자 구분','회원 등급','주문건수','매출','객단가','최근주문일']].sort_values('매출',ascending=False).reset_index(drop=True)
    search=st.text_input("🔍 검색 (아이디, 주문자명, 상호명)",key=f"{kp}_org_search" if kp else "org_search_main")
    if search: ba=ba[ba.apply(lambda r:search.lower() in str(r).lower(),axis=1)]
    st.dataframe(ba.style.format({'매출':'{:,.0f}원','주문건수':'{:,.0f}건','객단가':'{:,.0f}원'}),use_container_width=True,height=550)

def render_product_pareto(kp=""):
    st.markdown("#### 상품별 매출 TOP 20 (파레토 차트)")
    pa = filtered.groupby(['상품명','상품 코드']).agg(매출=('판매합계금액','sum'),수량=('주문 수량','sum'),주문건수=('주문 ID','nunique')).reset_index().sort_values('매출',ascending=False)
    top20=pa.head(20).copy(); ttl=pa['매출'].sum(); top20['누적비중']=(top20['매출'].cumsum()/ttl*100).round(1)
    fig=make_subplots(specs=[[{"secondary_y":True}]])
    fig.add_trace(go.Bar(x=[f"{i+1}. {n[:16]}" for i,n in enumerate(top20['상품명'])],y=top20['매출'],name='매출액',marker_color='#3366CC',opacity=0.8,hovertemplate='%{customdata[0]}<br>매출: %{customdata[1]}<extra></extra>',customdata=list(zip(top20['상품명'],[f"{v:,.0f}원" for v in top20['매출']]))),secondary_y=False)
    fig.add_trace(go.Scatter(x=[f"{i+1}. {n[:16]}" for i,n in enumerate(top20['상품명'])],y=top20['누적비중'],name='누적 비중',line=dict(color='#E8853D',width=3),mode='lines+markers',marker=dict(size=7),hovertemplate='%{customdata}<br>누적 비중: %{y:.1f}%<extra></extra>',customdata=top20['상품명'].tolist()),secondary_y=True)
    tvals3,ttexts3=krw_tickvals(top20['매출'])
    fig.update_layout(height=540,margin=dict(l=80,r=60,t=50,b=150),legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=12)),xaxis=dict(tickangle=45,tickfont=dict(size=9)))
    fig.update_yaxes(title_text="매출액",tickvals=tvals3,ticktext=ttexts3,tickfont=dict(size=11),secondary_y=False)
    fig.update_yaxes(title_text="누적 비중 (%)",range=[0,100],ticksuffix='%',tickfont=dict(size=11),secondary_y=True)
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"pareto"))

def render_product_sales_table(kp=""):
    st.markdown("#### 전체 상품 매출 현황")
    pa=filtered.groupby(['상품명','상품 코드']).agg(매출=('판매합계금액','sum'),수량=('주문 수량','sum'),주문건수=('주문 ID','nunique')).reset_index().sort_values('매출',ascending=False)
    sp=st.text_input("🔍 상품명/코드 검색",key=f"{kp}_prod_search" if kp else "prod_search_main")
    dp=pa.copy().reset_index(drop=True)
    if sp: dp=dp[dp.apply(lambda r:sp.lower() in str(r).lower(),axis=1)]
    st.dataframe(dp.style.format({'매출':'{:,.0f}원','수량':'{:,.0f}','주문건수':'{:,.0f}건'}),use_container_width=True,height=450)

def render_product_type_cross(kp=""):
    st.markdown("#### 회원구분별 × 상품 매출 크로스 (TOP 20)")
    pa=filtered.groupby(['상품명','상품 코드']).agg(매출=('판매합계금액','sum')).reset_index().sort_values('매출',ascending=False)
    t20n=pa.head(20)['상품명'].tolist()
    cp=filtered[filtered['상품명'].isin(t20n)].pivot_table(index='상품명',columns='주문자 구분',values='판매합계금액',aggfunc='sum',fill_value=0)
    cp['합계']=cp.sum(axis=1); cp=cp.sort_values('합계',ascending=False)
    st.dataframe(cp.style.format('{:,.0f}원'),use_container_width=True,height=550)

def render_product_monthly_trend(kp=""):
    st.markdown("#### 월별 상품 매출 추이")
    pa=filtered.groupby(['상품명','상품 코드']).agg(매출=('판매합계금액','sum')).reset_index().sort_values('매출',ascending=False)
    t5=pa.head(5)['상품명'].tolist()
    sel=st.multiselect("상품 선택",pa['상품명'].tolist(),default=t5,key=f"{kp}_prod_trend" if kp else "prod_trend_main")
    if sel:
        td=filtered[filtered['상품명'].isin(sel)].groupby(['주문월','상품명'])['판매합계금액'].sum().reset_index(); td['주문월_kr']=ym_series_kr(td['주문월'])
        fig=px.line(td,x='주문월_kr',y='판매합계금액',color='상품명',markers=True,color_discrete_sequence=COLORS)
        for tr in fig.data:
            tr.customdata=[f"{v:,.0f}원" for v in tr.y]; tr.hovertemplate='%{x}<br>'+(tr.name[:20]+'...' if len(tr.name)>20 else tr.name)+'<br>매출: %{customdata}<extra></extra>'
            if len(tr.name)>22: tr.name=tr.name[:22]+'...'
        fig.update_layout(height=480,margin=dict(l=70,r=30,t=30,b=120),legend=dict(orientation="h",yanchor="top",y=-0.15,x=0,font=dict(size=10)))
        st.plotly_chart(fig, use_container_width=True, key=_k(kp,"prod_trend_chart"))

def render_new_member_trend(kp=""):
    st.markdown("#### 월별 신규가입자 추이 (회원타입별)")
    jm=filtered_members.groupby(['가입월','회원타입']).size().reset_index(name='가입자수'); jm['가입월_kr']=ym_series_kr(jm['가입월'])
    fig=px.bar(jm,x='가입월_kr',y='가입자수',color='회원타입',color_discrete_sequence=COLORS)
    for tr in fig.data: tr.hovertemplate='%{x}<br>'+tr.name+': %{y:,}처<extra></extra>'
    fig.update_layout(height=480,barmode='stack',margin=dict(l=60,r=30,t=50,b=70),legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=11)))
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"new_member"))

def render_grade_member_bar(kp=""):
    st.markdown("#### 회원등급별 가입자 분포")
    gd=filtered_members['회원등급'].value_counts().reset_index(); gd.columns=['등급','수']
    fig=px.bar(gd,x='수',y='등급',orientation='h',color_discrete_sequence=COLORS)
    fig.update_traces(text=[fmt_num(v) for v in gd['수']],textposition='outside',textfont=dict(size=11),hovertemplate='%{y}: %{x:,}처<extra></extra>')
    fig.update_layout(height=max(420,len(gd)*35+140),margin=dict(l=140,r=80,t=30,b=40),showlegend=False)
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"grade_member"))

def render_first_order_days(kp=""):
    st.markdown("#### 가입 후 첫 주문까지 소요일")
    mo_df=orders.groupby('주문자 ID').agg(첫주문일=('주문일','min')).reset_index()
    mg=filtered_members.merge(mo_df,left_on='아이디',right_on='주문자 ID',how='inner'); mg['소요일']=(mg['첫주문일']-mg['가입일']).dt.days; mg=mg[mg['소요일']>=0]
    bins=[0,1,8,15,31,61,91,9999]; lb=['당일','1~7일','8~14일','15~30일','31~60일','61~90일','90일+']
    mg['구간']=pd.cut(mg['소요일'],bins=bins,labels=lb,right=False); dh=mg['구간'].value_counts().reindex(lb).fillna(0).reset_index(); dh.columns=['구간','회원수']
    fig=px.bar(dh,x='구간',y='회원수',color_discrete_sequence=['#3366CC'])
    fig.update_traces(text=[fmt_num(v) for v in dh['회원수']],textposition='outside',textfont=dict(size=11),hovertemplate='%{x}: %{y:,}처<extra></extra>')
    fig.update_layout(height=450,margin=dict(l=60,r=30,t=30,b=60),showlegend=False)
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"first_order"))

def render_order_count_dist(kp=""):
    st.markdown("#### 주문횟수 구간별 회원 분포")
    mo_df=orders.groupby('주문자 ID').agg(주문건수=('주문 ID','nunique')).reset_index()
    bo=[1,2,4,6,11,21,9999]; lo=['1회','2~3회','4~5회','6~10회','11~20회','20회+']
    mo_df['구간']=pd.cut(mo_df['주문건수'],bins=bo,labels=lo,right=False); od=mo_df['구간'].value_counts().reindex(lo).fillna(0).reset_index(); od.columns=['구간','회원수']
    fig=px.bar(od,x='구간',y='회원수',color_discrete_sequence=COLORS)
    fig.update_traces(text=[fmt_num(v) for v in od['회원수']],textposition='outside',textfont=dict(size=11),hovertemplate='%{x}: %{y:,}처<extra></extra>')
    fig.update_layout(height=450,margin=dict(l=60,r=30,t=30,b=60),showlegend=False)
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"order_dist"))

def render_cohort_heatmap(kp=""):
    st.markdown("#### 코호트 리텐션 히트맵 (가입월 × 경과월별 재구매율)")
    cm=members[members['가입일'].notna()].copy(); cm['코호트']=cm['가입일'].dt.to_period('M').astype(str)
    om=orders.groupby('주문자 ID')['주문월'].apply(set).to_dict(); cohorts=sorted(cm['코호트'].unique()); mx=12; rd=[]
    for c in cohorts:
        cu=cm[cm['코호트']==c]['아이디'].tolist(); sz=len(cu)
        if sz==0: continue
        row={'코호트':c,'크기':sz}; cp=pd.Period(c,freq='M')
        for o in range(mx):
            t=str(cp+o); a=sum(1 for u in cu if t in om.get(u,set())); row[f'{o}개월']=round(a/sz*100,1)
        rd.append(row)
    if rd:
        rdf=pd.DataFrame(rd); mc=[f'{o}개월' for o in range(mx)]; zv=rdf[mc].values
        hover_texts=[]
        for _,r in rdf.iterrows():
            cp2=pd.Period(r['코호트'],freq='M')
            hover_texts.append([f"{o}개월 후 ({(cp2+o).year}년 {(cp2+o).month}월)" for o in range(mx)])
        fig=go.Figure(data=go.Heatmap(z=zv,x=mc,y=[f"{to_ym_kr(r['코호트'])} ({r['크기']}처)" for _,r in rdf.iterrows()],colorscale=[[0,'#F0F2F5'],[0.3,'#A8D5A2'],[1,'#27AE60']],text=[[f'{v:.1f}%' if v>0 else '-' for v in row] for row in zv],texttemplate='%{text}',textfont=dict(size=10),customdata=hover_texts,hovertemplate='%{y}<br>%{customdata}: %{z:.1f}%<extra></extra>'))
        fig.update_layout(height=max(400,len(rd)*32+140),margin=dict(l=190,r=20,t=20,b=50),yaxis=dict(tickfont=dict(size=10),autorange="reversed"),xaxis=dict(tickfont=dict(size=11)))
        st.plotly_chart(fig, use_container_width=True, key=_k(kp,"cohort"))

def render_dormant_analysis(kp=""):
    st.markdown("#### 휴면 · 활성 회원 분석")
    st.markdown("""<div style='background:#f8fafc;border-radius:10px;padding:12px 20px;margin-bottom:16px;border:1px solid #e2e8f0;font-size:0.85rem;color:#475569;'>
    🟢 <b>활성</b> 최근 3개월 이내 주문 &nbsp;|&nbsp; 🟡 <b>단기휴면</b> 3~6개월 &nbsp;|&nbsp; 🟠 <b>중기휴면</b> 6~12개월 &nbsp;|&nbsp; 🔴 <b>장기휴면</b> 12개월 이상 &nbsp;|&nbsp; ⚪ <b>미구매</b> 가입 후 주문 없음
    </div>""", unsafe_allow_html=True)
    base_date=orders['주문일'].max()
    last_order=orders.groupby('주문자 ID')['주문일'].max().reset_index(); last_order.columns=['아이디','마지막주문일']
    dormant_df=members.copy(); dormant_df=dormant_df.merge(last_order,on='아이디',how='left')
    dormant_df['휴면경과일']=(base_date-dormant_df['마지막주문일']).dt.days
    def classify_status(row):
        if pd.isna(row['마지막주문일']): return '미구매'
        d=row['휴면경과일']
        if d<=90: return '활성'
        elif d<=180: return '단기휴면'
        elif d<=365: return '중기휴면'
        else: return '장기휴면'
    dormant_df['활성구분']=dormant_df.apply(classify_status,axis=1)
    status_order=['활성','단기휴면','중기휴면','장기휴면','미구매']
    status_colors={'활성':'#27AE60','단기휴면':'#F1C40F','중기휴면':'#E8853D','장기휴면':'#E74C3C','미구매':'#BDC3C7'}
    status_counts=dormant_df['활성구분'].value_counts().reindex(status_order).fillna(0).reset_index(); status_counts.columns=['구분','수']
    c1,c2=st.columns(2)
    with c1:
        st.markdown("##### 회원 활성 현황")
        fig=make_donut(status_counts,'구분','수',colors=[status_colors[s] for s in status_counts['구분']],value_label='회원수',unit='처')
        fig.update_layout(height=450); fig.layout.annotations[0].text=f"<b>전체</b><br>{fmt_num(len(dormant_df))}처"
        st.plotly_chart(fig, use_container_width=True, key=_k(kp,"dormant_donut"))
    with c2:
        st.markdown("##### 구간별 회원 수")
        fig=go.Figure()
        fig.add_trace(go.Bar(x=status_counts['구분'],y=status_counts['수'],marker_color=[status_colors[s] for s in status_counts['구분']],text=[fmt_num(v) for v in status_counts['수']],textposition='outside',textfont=dict(size=12),hovertemplate='%{x}: %{y:,}처<extra></extra>'))
        fig.update_layout(height=450,showlegend=False,margin=dict(l=60,r=30,t=30,b=40),xaxis=dict(title='',tickfont=dict(size=12),categoryorder='array',categoryarray=status_order),yaxis=dict(title='회원 수 (처)',tickfont=dict(size=11)))
        st.plotly_chart(fig, use_container_width=True, key=_k(kp,"dormant_bar"))

def _build_referral_data():
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
    return pd.DataFrame(ra.values())[['추천인','유형','추천인코드','피추천인수','피추천인매출']]

def render_referral_count_bar(kp=""):
    st.markdown("#### 추천인 유형별 피추천인 수")
    rdf_local=_build_referral_data()
    tc={'영업팀':'#3366CC','대리점':'#E8853D','케어포':'#27AE60'}
    tr_df=rdf_local.groupby('유형')['피추천인수'].sum().reset_index()
    fig=px.bar(tr_df,x='유형',y='피추천인수',color='유형',color_discrete_map=tc)
    fig.update_traces(text=[fmt_num(v) for v in tr_df['피추천인수']],textposition='outside',textfont=dict(size=12),hovertemplate='%{x}: %{y:,}회원<extra></extra>')
    fig.update_layout(height=450,showlegend=False,margin=dict(l=60,r=30,t=30,b=40))
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"ref_count"))

def render_referral_sales_donut(kp=""):
    st.markdown("#### 추천인 유형별 피추천인 매출")
    rdf_local=_build_referral_data()
    tc={'영업팀':'#3366CC','대리점':'#E8853D','케어포':'#27AE60'}
    ts_ref=rdf_local.groupby('유형')['피추천인매출'].sum().reset_index()
    fig=make_donut(ts_ref,'유형','피추천인매출',colors=[tc.get(t,'#999') for t in ts_ref['유형']])
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"ref_donut"))

def render_referral_table(kp=""):
    st.markdown("#### 추천인별 현황")
    rdf_local=_build_referral_data().sort_values('피추천인매출',ascending=False).reset_index(drop=True)
    sr=st.text_input("🔍 추천인 검색",key=f"{kp}_ref_table_search" if kp else "ref_table_search_main")
    if sr: rdf_local=rdf_local[rdf_local.apply(lambda r:sr.lower() in str(r).lower(),axis=1)]
    st.dataframe(rdf_local.style.format({'피추천인수':'{:,.0f}','피추천인매출':'{:,.0f}원'}),use_container_width=True,height=550)

def render_dealer_commission(kp=""):
    st.markdown("#### 대리점 피추천인 매출 및 판매수수료 집계")
    st.caption("구매확정일자 기준 · 전월 26일~당월 25일 · 상품별 수수료율 적용")

    if dealer_raw.empty:
        st.info("대리점 피추천인 주문내역 데이터가 없습니다.")
        return

    dr = dealer_raw.copy()
    dr.columns = [c.replace('\n','').strip() for c in dr.columns]

    def parse_rate(val):
        if pd.isna(val): return 0
        s = str(val).replace('%','').strip()
        try: return float(s) / 100
        except: return 0

    def parse_amount(val):
        if pd.isna(val): return 0
        if isinstance(val, (int, float)): return float(val)
        try: return float(str(val).replace(',','').strip())
        except: return 0

    dr['판매금액_num'] = dr['판매금액'].apply(parse_amount)
    dr['수수료율_num'] = dr['상품 수수료율'].apply(parse_rate)
    dr['수수료금액'] = dr['판매금액_num'] * dr['수수료율_num']
    dr['정산월_key'] = dr['정산연도'].astype(str).str.zfill(4) + '-' + dr['정산월'].astype(str).str.zfill(2)

    dealer_opts = sorted(dr['추천인명(상호명)'].unique().tolist())
    sel_dealers = st.multiselect("대리점 필터", dealer_opts, default=[], 
        placeholder="전체 (선택 안 하면 전체)", key=_k(kp,"dealer_filter") if kp else "dealer_filter_main")
    dr_chart = dr[dr['추천인명(상호명)'].isin(sel_dealers)] if sel_dealers else dr

    monthly = dr_chart.groupby('정산월_key').agg(
        피추천인매출=('판매금액_num','sum'),
        판매수수료=('수수료금액','sum')
    ).reset_index().sort_values('정산월_key')
    monthly['정산월_kr'] = ym_series_kr(monthly['정산월_key'])

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(
        x=monthly['정산월_kr'], y=monthly['피추천인매출'],
        name='피추천인 매출', marker_color='#3366CC', opacity=0.8,
        hovertemplate='%{x}<br>피추천인 매출: %{customdata}<extra></extra>',
        customdata=[f"{v:,.0f}원" for v in monthly['피추천인매출']]
    ), secondary_y=False)
    fig.add_trace(go.Scatter(
        x=monthly['정산월_kr'], y=monthly['판매수수료'],
        name='판매수수료', line=dict(color='#E74C3C', width=3),
        mode='lines+markers+text', marker=dict(size=8),
        text=[fmt_krw_short(v) for v in monthly['판매수수료']],
        textposition='top center', textfont=dict(size=11, color='#E74C3C'),
        hovertemplate='%{x}<br>판매수수료: %{customdata}<extra></extra>',
        customdata=[f"{v:,.0f}원" for v in monthly['판매수수료']]
    ), secondary_y=True)
    tvals, ttexts = krw_tickvals(monthly['피추천인매출'])
    tvals2, ttexts2 = krw_tickvals(monthly['판매수수료'])
    fig.update_layout(height=480, margin=dict(l=80,r=60,t=50,b=70),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=12)),
        xaxis=dict(tickfont=dict(size=12)))
    fig.update_yaxes(title_text="피추천인 매출", tickvals=tvals, ticktext=ttexts, tickfont=dict(size=11), secondary_y=False)
    fig.update_yaxes(title_text="판매수수료", tickvals=tvals2, ticktext=ttexts2, tickfont=dict(size=11), secondary_y=True)
    st.markdown("##### 월별 피추천인 매출 · 판매수수료 추이")
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"dealer_chart"))

    st.markdown("##### 대리점별 × 정산월 상세")
    dealer_monthly = dr_chart.groupby(['추천인명(상호명)','정산월_key']).agg(
        피추천인매출=('판매금액_num','sum'),
        판매수수료=('수수료금액','sum')
    ).reset_index()
    pivot_sales = dealer_monthly.pivot(index='추천인명(상호명)', columns='정산월_key', values='피추천인매출').fillna(0)
    pivot_comm = dealer_monthly.pivot(index='추천인명(상호명)', columns='정산월_key', values='판매수수료').fillna(0)
    pivot_sales.columns = pd.MultiIndex.from_tuples([(to_ym_kr(c),'피추천인매출') for c in pivot_sales.columns])
    pivot_comm.columns = pd.MultiIndex.from_tuples([(to_ym_kr(c),'판매수수료') for c in pivot_comm.columns])
    combined = pd.concat([pivot_sales, pivot_comm], axis=1).sort_index(axis=1)
    combined.loc['합계'] = combined.sum()
    st.dataframe(combined.style.format('{:,.0f}원'), use_container_width=True,
        height=550)

def render_dealer_commission_forecast(kp=""):
    st.markdown("#### 🔮 대리점 판매수수료 예측 (당월)")
    st.caption("주문내역 기준 · 구매확정일 있는 주문 + 미확정 주문(주문일+14일 예상) · 수수료율 10% 고정")

    # 1단계: 대리점 추천인 목록
    dealer_refs = referrals_df[referrals_df['회원그룹'] == '대리점 회원'].copy()
    if dealer_refs.empty:
        st.info("대리점 회원 추천인 데이터가 없습니다.")
        return

    # 2단계: 피추천인 사업자번호 → 주문자 ID
    b2u = members.set_index('사업자번호')['아이디'].to_dict()
    dealer_refs['피추천인_아이디'] = dealer_refs['피추천인 사업자 번호'].map(b2u)
    dealer_map = dealer_refs.dropna(subset=['피추천인_아이디']).groupby('추천인')['피추천인_아이디'].apply(set).to_dict()

    if not dealer_map:
        st.info("매칭된 피추천인 데이터가 없습니다.")
        return

    # 전체 피추천인 아이디 목록
    all_buyer_ids = set().union(*dealer_map.values())

    # 3단계: 주문 필터링
    # 구매확정일 있는 주문 → 정산월 그대로
    # 구매확정일 없는 주문 → 주문일 + 14일로 예상 정산월 계산
    orders_target = orders[orders['주문자 ID'].isin(all_buyer_ids)].copy()

    def calc_settlement_month(row):
        if pd.notna(row.get('구매확정일')) and row['정산월'] is not None:
            return row['정산월']
        # 미확정: 주문일 + 14일
        expected = row['주문일'] + pd.Timedelta(days=14)
        if pd.isna(expected): return None
        if expected.day >= 26:
            if expected.month == 12: return f"{expected.year+1}-01"
            else: return f"{expected.year}-{str(expected.month+1).zfill(2)}"
        else:
            return f"{expected.year}-{str(expected.month).zfill(2)}"

    orders_target['예상정산월'] = orders_target.apply(calc_settlement_month, axis=1)

    # 당월 계산
    today = pd.Timestamp.now()
    if today.day >= 26:
        if today.month == 12: current_month = f"{today.year+1}-01"
        else: current_month = f"{today.year}-{str(today.month+1).zfill(2)}"
    else:
        current_month = f"{today.year}-{str(today.month).zfill(2)}"

    # 당월 주문만 필터
    orders_curr = orders_target[orders_target['예상정산월'] == current_month].copy()

    if orders_curr.empty:
        st.info(f"{to_ym_kr(current_month)} 예측 데이터가 없습니다.")
        return

    st.markdown(f"**기준 정산월: {to_ym_kr(current_month)}**")

    # 확정/미확정 구분
    orders_curr['확정여부'] = orders_curr['정산월'].apply(
        lambda x: '확정' if x == current_month else '미확정(예측)'
    )

    # 4단계: 대리점별 집계
    # 주문자 ID → 대리점 역매핑
    id_to_dealer = {}
    for dealer, buyer_ids in dealer_map.items():
        for bid in buyer_ids:
            id_to_dealer[bid] = dealer

    orders_curr['대리점'] = orders_curr['주문자 ID'].map(id_to_dealer)
    orders_curr = orders_curr.dropna(subset=['대리점'])

    # 대리점별 × 확정여부 집계
    summary = orders_curr.groupby(['대리점','확정여부']).agg(
        피추천인매출=('판매합계금액','sum')
    ).reset_index()
    summary['판매수수료'] = summary['피추천인매출'] * 0.1

    # 피벗
    pivot = summary.pivot_table(
        index='대리점', columns='확정여부',
        values=['피추천인매출','판매수수료'], aggfunc='sum', fill_value=0
    )
    pivot.columns = [f"{b}_{a}" for a,b in pivot.columns]
    
    # 합계 컬럼 추가
    if '확정_피추천인매출' not in pivot.columns: pivot['확정_피추천인매출'] = 0
    if '미확정(예측)_피추천인매출' not in pivot.columns: pivot['미확정(예측)_피추천인매출'] = 0
    if '확정_판매수수료' not in pivot.columns: pivot['확정_판매수수료'] = 0
    if '미확정(예측)_판매수수료' not in pivot.columns: pivot['미확정(예측)_판매수수료'] = 0

    pivot['예상총매출'] = pivot['확정_피추천인매출'] + pivot['미확정(예측)_피추천인매출']
    pivot['예상총수수료'] = pivot['확정_판매수수료'] + pivot['미확정(예측)_판매수수료']
    pivot = pivot.sort_values('예상총수수료', ascending=False)
    pivot.loc['합계'] = pivot.sum()

    # KPI
    total_confirmed = pivot.loc['합계','확정_판매수수료'] if '합계' in pivot.index else 0
    total_expected = pivot.loc['합계','미확정(예측)_판매수수료'] if '합계' in pivot.index else 0
    total_forecast = pivot.loc['합계','예상총수수료'] if '합계' in pivot.index else 0

    c1, c2, c3 = st.columns(3)
    c1.markdown(kpi_card("확정 수수료", fmt_krw_short(total_confirmed), "원"), unsafe_allow_html=True)
    c2.markdown(kpi_card("예측 수수료", fmt_krw_short(total_expected), "원"), unsafe_allow_html=True)
    c3.markdown(kpi_card("예상 합계", fmt_krw_short(total_forecast), "원"), unsafe_allow_html=True)

    # 테이블
    display_cols = ['확정_피추천인매출','확정_판매수수료','미확정(예측)_피추천인매출','미확정(예측)_판매수수료','예상총매출','예상총수수료']
    display_cols = [c for c in display_cols if c in pivot.columns]
    st.dataframe(
        pivot[display_cols].style.format('{:,.0f}원'),
        use_container_width=True, height=550
    )

def render_carefor_grade_sales(kp=""):
    st.markdown("#### 케어포 등급별 매출")
    cfg=['케어포-시설','케어포-공생','케어포-주야간','케어포-방문','케어포-일반','케어포-종사자','케어포-보호자']
    co=filtered[filtered['회원 등급'].isin(cfg)]; grade_order=['시설','공생','주야간','방문','일반','보호자','종사자']
    cga=co.groupby('회원 등급').agg(매출=('판매합계금액','sum')).reset_index(); cga['등급']=cga['회원 등급'].str.replace('케어포-','')
    cga['등급']=pd.Categorical(cga['등급'],categories=grade_order,ordered=True); cga=cga.sort_values('등급')
    fig=px.bar(cga,x='등급',y='매출',color_discrete_sequence=['#3366CC'])
    fig.update_traces(text=[fmt_krw_short(v) for v in cga['매출']],textposition='outside',textfont=dict(size=11),hovertemplate='%{x}<br>매출: %{customdata}<extra></extra>',customdata=[f"{v:,.0f}원" for v in cga['매출']])
    fig.update_layout(height=420,margin=dict(l=60,r=20,t=30,b=40),showlegend=False,xaxis=dict(title='',tickfont=dict(size=12),categoryorder='array',categoryarray=grade_order))
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"cf_grade_sales"))

def render_carefor_grade_orders(kp=""):
    st.markdown("#### 케어포 등급별 주문건수")
    cfg=['케어포-시설','케어포-공생','케어포-주야간','케어포-방문','케어포-일반','케어포-종사자','케어포-보호자']
    co=filtered[filtered['회원 등급'].isin(cfg)]; grade_order=['시설','공생','주야간','방문','일반','보호자','종사자']
    cga=co.groupby('회원 등급').agg(주문건수=('주문 ID','nunique')).reset_index(); cga['등급']=cga['회원 등급'].str.replace('케어포-','')
    cga['등급']=pd.Categorical(cga['등급'],categories=grade_order,ordered=True); cga=cga.sort_values('등급')
    fig=px.bar(cga,x='등급',y='주문건수',color_discrete_sequence=['#E8853D'])
    fig.update_traces(text=[fmt_num(v) for v in cga['주문건수']],textposition='outside',textfont=dict(size=11),hovertemplate='%{x}<br>주문: %{y:,}건<extra></extra>')
    fig.update_layout(height=420,margin=dict(l=60,r=20,t=30,b=40),showlegend=False,xaxis=dict(title='',tickfont=dict(size=12),categoryorder='array',categoryarray=grade_order))
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"cf_grade_orders"))

def render_carefor_monthly_trend(kp=""):
    st.markdown("#### 케어포 월별 매출 추이")
    cfg=['케어포-시설','케어포-공생','케어포-주야간','케어포-방문','케어포-일반','케어포-종사자','케어포-보호자']
    co=filtered[filtered['회원 등급'].isin(cfg)]
    cf_monthly=co.groupby('주문월')['판매합계금액'].sum().reset_index(); cf_monthly['주문월_kr']=ym_series_kr(cf_monthly['주문월'])
    tvals_cfm,ttexts_cfm=krw_tickvals(cf_monthly['판매합계금액'])
    fig=go.Figure()
    fig.add_trace(go.Bar(x=cf_monthly['주문월_kr'],y=cf_monthly['판매합계금액'],name='매출액',marker_color='#27AE60',opacity=0.8,text=[fmt_krw_short(v) for v in cf_monthly['판매합계금액']],textposition='outside',textfont=dict(size=10),hovertemplate='%{x}<br>매출: %{customdata}<extra></extra>',customdata=[f"{v:,.0f}원" for v in cf_monthly['판매합계금액']]))
    z=np.polyfit(range(len(cf_monthly)),cf_monthly['판매합계금액'].values,1); trend=np.polyval(z,range(len(cf_monthly)))
    fig.add_trace(go.Scatter(x=cf_monthly['주문월_kr'],y=trend,name='추세선',line=dict(color='#E74C3C',width=2,dash='dash'),mode='lines',hoverinfo='skip'))
    fig.update_layout(height=450,margin=dict(l=70,r=30,t=30,b=60),legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=11)),xaxis=dict(title='',tickfont=dict(size=12)),yaxis=dict(title='매출액',tickvals=tvals_cfm,ticktext=ttexts_cfm,tickfont=dict(size=11)))
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"cf_monthly"))

def render_carefor_new_member_trend(kp=""):
    st.markdown("#### 케어포 등급별 신규가입 추이")
    cf_filtered_local=filtered_members[filtered_members['회원타입']=='케어포']
    cj=cf_filtered_local.groupby(['가입월','회원등급']).size().reset_index(name='가입자수'); cj['가입월_kr']=ym_series_kr(cj['가입월'])
    ccl={'케어포-시설':'#3366CC','케어포-공생':'#E8853D','케어포-주야간':'#27AE60','케어포-방문':'#9B59B6','케어포-일반':'#F39C12','케어포-종사자':'#E74C3C','케어포-보호자':'#1ABC9C'}
    fig=px.bar(cj,x='가입월_kr',y='가입자수',color='회원등급',color_discrete_map=ccl)
    for tr in fig.data: tr.hovertemplate='%{x}<br>'+tr.name+': %{y:,}처<extra></extra>'
    fig.update_layout(height=480,barmode='stack',margin=dict(l=60,r=20,t=30,b=100),legend=dict(orientation="h",yanchor="top",y=-0.12,x=0,font=dict(size=10)))
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"cf_new_member"))

def render_pnl_waterfall(kp=""):
    st.markdown("#### 손익 워터폴")
    if len(bw_filtered)==0: st.warning("BW 데이터 없음"); return
    bw=bw_filtered.copy()
    rev=bw['I.매출액(FI기준)'].sum(); cogs=bw['II.매출원가'].sum(); gp=bw['III.매출총이익'].sum(); sga=bw['IV.판매비 및 관리비'].sum(); oi=bw['V.영업이익I'].sum()
    wf_values=[rev,-cogs,gp,-sga,oi]
    fig=go.Figure(go.Waterfall(x=['매출액','매출원가','매출총이익','판관비','영업이익'],y=wf_values,measure=['absolute','relative','total','relative','total'],connector=dict(line=dict(color="#94a3b8",width=1)),increasing=dict(marker=dict(color='#27AE60')),decreasing=dict(marker=dict(color='#E74C3C')),totals=dict(marker=dict(color='#3366CC')),text=[fmt_krw_short(abs(v)) for v in wf_values],textposition='outside',textfont=dict(size=12),hovertemplate='%{x}<br>금액: %{customdata}<extra></extra>',customdata=[f"{abs(v):,.0f}원" for v in wf_values]))
    wf_tvals,wf_ttexts=krw_tickvals(pd.Series([abs(v) for v in wf_values]))
    fig.update_layout(height=480,margin=dict(l=80,r=30,t=50,b=60),showlegend=False,xaxis=dict(tickfont=dict(size=13)),yaxis=dict(title='금액',tickvals=wf_tvals,ticktext=wf_ttexts,tickfont=dict(size=11)))
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"waterfall"))

def render_pnl_monthly_trend(kp=""):
    st.markdown("#### 월별 손익 추이")
    if len(bw_filtered)==0: st.warning("BW 데이터 없음"); return
    bw=bw_filtered.copy()
    bw_monthly=bw.groupby('연월').agg(매출액=('I.매출액(FI기준)','sum'),매출총이익=('III.매출총이익','sum'),영업이익=('V.영업이익I','sum')).reset_index().sort_values('연월')
    bw_monthly['영업이익률']=np.where(bw_monthly['매출액']!=0,bw_monthly['영업이익']/bw_monthly['매출액']*100,0)
    bw_monthly['연월_kr']=ym_series_kr(bw_monthly['연월'])
    fig=make_subplots(specs=[[{"secondary_y":True}]])
    for col_name,color,name in [('매출액','#3366CC','매출액'),('매출총이익','#27AE60','매출총이익'),('영업이익','#E8853D','영업이익')]:
        fig.add_trace(go.Bar(x=bw_monthly['연월_kr'],y=bw_monthly[col_name],name=name,marker_color=color,opacity=0.85,hovertemplate='%{x}<br>'+name+': %{customdata}<extra></extra>',customdata=[f"{v:,.0f}원" for v in bw_monthly[col_name]]),secondary_y=False)
    fig.add_trace(go.Scatter(x=bw_monthly['연월_kr'],y=bw_monthly['영업이익률'],name='영업이익률',line=dict(color='#E74C3C',width=3),mode='lines+markers+text',marker=dict(size=8),text=[f"{v:.1f}%" for v in bw_monthly['영업이익률']],textposition='top center',textfont=dict(size=11,color='#E74C3C'),hovertemplate='%{x}<br>영업이익률: %{y:.1f}%<extra></extra>'),secondary_y=True)
    fig.update_layout(height=500,barmode='group',margin=dict(l=80,r=60,t=50,b=70),legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=11)),xaxis=dict(tickfont=dict(size=12)))
    monthly_tvals,monthly_ttexts=krw_tickvals(bw_monthly[['매출액','매출총이익','영업이익']].abs().max())
    fig.update_yaxes(title_text="금액",tickvals=monthly_tvals,ticktext=monthly_ttexts,tickfont=dict(size=11),secondary_y=False)
    fig.update_yaxes(title_text="영업이익률 (%)",tickfont=dict(size=11),ticksuffix='%',secondary_y=True)
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"pnl_monthly"))

def render_channel_pnl(kp=""):
    st.markdown("#### 채널별 손익 비교")
    if len(bw_filtered)==0: st.warning("BW 데이터 없음"); return
    bw=bw_filtered.copy()
    ch_pnl=bw.groupby('채널').agg(매출액=('I.매출액(FI기준)','sum'),영업이익=('V.영업이익I','sum')).reset_index()
    ch_pnl['영업이익률']=np.where(ch_pnl['매출액']!=0,ch_pnl['영업이익']/ch_pnl['매출액']*100,0); ch_pnl=ch_pnl.sort_values('매출액',ascending=True)
    fig=go.Figure()
    fig.add_trace(go.Bar(x=ch_pnl['매출액'],y=ch_pnl['채널'],name='매출액',orientation='h',marker_color='#3366CC',opacity=0.8,text=[fmt_krw_short(v) for v in ch_pnl['매출액']],textposition='outside',textfont=dict(size=10),hovertemplate='%{y}<br>매출액: %{customdata[0]}<br>영업이익률: %{customdata[1]}<extra></extra>',customdata=list(zip([f"{v:,.0f}원" for v in ch_pnl['매출액']],[f"{v:.1f}%" for v in ch_pnl['영업이익률']]))))
    fig.add_trace(go.Bar(x=ch_pnl['영업이익'],y=ch_pnl['채널'],name='영업이익',orientation='h',marker_color='#E8853D',opacity=0.8,text=[fmt_krw_short(v) for v in ch_pnl['영업이익']],textposition='outside',textfont=dict(size=10),hovertemplate='%{y}<br>영업이익: %{customdata[0]}<br>영업이익률: %{customdata[1]}<extra></extra>',customdata=list(zip([f"{v:,.0f}원" for v in ch_pnl['영업이익']],[f"{v:.1f}%" for v in ch_pnl['영업이익률']]))))
    ch_tvals,ch_ttexts=krw_tickvals(ch_pnl['매출액'])
    fig.update_layout(height=max(450,len(ch_pnl)*50+140),barmode='group',margin=dict(l=130,r=100,t=30,b=40),legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=11)),xaxis=dict(title='금액',tickvals=ch_tvals,ticktext=ch_ttexts,tickfont=dict(size=11)))
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"ch_pnl"))

def render_channel_pnl_table(kp=""):
    st.markdown("#### 채널별 손익 상세")
    if len(bw_filtered)==0: st.warning("BW 데이터 없음"); return
    bw=bw_filtered.copy()
    ch_pnl=bw.groupby('채널').agg(매출액=('I.매출액(FI기준)','sum'),매출원가=('II.매출원가','sum'),매출총이익=('III.매출총이익','sum'),판관비=('IV.판매비 및 관리비','sum'),영업이익=('V.영업이익I','sum')).reset_index()
    ch_pnl['매출총이익률']=np.where(ch_pnl['매출액']!=0,ch_pnl['매출총이익']/ch_pnl['매출액']*100,0)
    ch_pnl['영업이익률']=np.where(ch_pnl['매출액']!=0,ch_pnl['영업이익']/ch_pnl['매출액']*100,0)
    ch_display=ch_pnl.sort_values('매출액',ascending=False)[['채널','매출액','매출원가','매출총이익','매출총이익률','판관비','영업이익','영업이익률']].reset_index(drop=True)
    st.dataframe(ch_display.style.format({'매출액':'{:,.0f}원','매출원가':'{:,.0f}원','매출총이익':'{:,.0f}원','매출총이익률':'{:.1f}%','판관비':'{:,.0f}원','영업이익':'{:,.0f}원','영업이익률':'{:.1f}%'}),use_container_width=True,height=450)

def render_sga_composition(kp=""):
    st.markdown("#### 판관비 구성 분석")
    if len(bw_filtered)==0: st.warning("BW 데이터 없음"); return
    bw=bw_filtered.copy()
    sga=bw['IV.판매비 및 관리비'].sum(); adv=bw['IV.6.광고선전비'].sum(); freight=bw['IV.7.운반비'].sum(); commission=bw['IV.8.판매수수료'].sum(); promo=bw['IV.9.판촉비'].sum(); etc_sga=sga-adv-freight-commission-promo
    sga_df=pd.DataFrame({'항목':['광고선전비','운반비','판매수수료','판촉비','기타판관비'],'금액':[adv,freight,commission,promo,etc_sga]})
    sga_df=sga_df[sga_df['금액']>0].sort_values('금액',ascending=False)
    fig=make_donut(sga_df,'항목','금액'); fig.update_layout(height=480)
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"sga_donut"))

def render_sga_monthly_trend(kp=""):
    st.markdown("#### 월별 판관비 구성 추이")
    if len(bw_filtered)==0: st.warning("BW 데이터 없음"); return
    bw=bw_filtered.copy()
    sga_monthly=bw.groupby('연월').agg(광고선전비=('IV.6.광고선전비','sum'),운반비=('IV.7.운반비','sum'),판매수수료=('IV.8.판매수수료','sum'),판촉비=('IV.9.판촉비','sum')).reset_index().sort_values('연월')
    sga_monthly['기타판관비']=bw.groupby('연월')['IV.판매비 및 관리비'].sum().values-sga_monthly[['광고선전비','운반비','판매수수료','판촉비']].sum(axis=1).values
    sga_monthly['연월_kr']=ym_series_kr(sga_monthly['연월'])
    sga_cols=['광고선전비','운반비','판매수수료','판촉비','기타판관비']; sga_colors=['#3366CC','#E8853D','#27AE60','#9B59B6','#94a3b8']
    fig=go.Figure()
    for col_name,color in zip(sga_cols,sga_colors):
        fig.add_trace(go.Bar(x=sga_monthly['연월_kr'],y=sga_monthly[col_name],name=col_name,marker_color=color,hovertemplate='%{x}<br>'+col_name+': %{customdata}<extra></extra>',customdata=[f"{v:,.0f}원" for v in sga_monthly[col_name]]))
    sga_tvals,sga_ttexts=krw_tickvals(sga_monthly[['광고선전비','운반비','판매수수료','판촉비','기타판관비']].sum(axis=1))
    fig.update_layout(height=480,barmode='stack',margin=dict(l=70,r=20,t=30,b=70),legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=10)),xaxis=dict(tickfont=dict(size=12)),yaxis=dict(title='판관비',tickvals=sga_tvals,ticktext=sga_ttexts,tickfont=dict(size=11)))
    st.plotly_chart(fig, use_container_width=True, key=_k(kp,"sga_monthly"))

def render_product_pnl_hierarchy(kp=""):
    st.markdown("#### 제품계층구조별 수익성 분석")
    if len(bw_filtered)==0: st.warning("BW 데이터 없음"); return
    bw=bw_filtered.copy()
    prod_ch_opts=sorted(bw['채널'].unique().tolist())
    sel_prod_ch=st.multiselect("채널 필터",prod_ch_opts,default=[],placeholder="전체",key=f"{kp}_prod_ch" if kp else "prod_ch_main")
    bw=bw[bw['채널'].isin(sel_prod_ch)] if sel_prod_ch else bw
    pfx=f"{kp}_ppnl" if kp else "ppnl_main"
    def _render_pnl(df,group_col,tab_key):
        pnl=df.groupby(group_col).agg(매출액=('I.매출액(FI기준)','sum'),매출총이익=('III.매출총이익','sum'),영업이익=('V.영업이익I','sum')).reset_index()
        pnl['매출총이익률']=np.where(pnl['매출액']!=0,pnl['매출총이익']/pnl['매출액']*100,0)
        pnl['영업이익률']=np.where(pnl['매출액']!=0,pnl['영업이익']/pnl['매출액']*100,0); pnl=pnl.sort_values('매출액',ascending=True)
        fig=go.Figure()
        fig.add_trace(go.Bar(x=pnl['매출액'],y=pnl[group_col],name='매출액',orientation='h',marker_color='#3366CC',opacity=0.8,text=[fmt_krw_short(v) for v in pnl['매출액']],textposition='outside',textfont=dict(size=10),hovertemplate='%{y}<br>매출액: %{customdata[0]}<br>영업이익률: %{customdata[1]}<extra></extra>',customdata=list(zip([f"{v:,.0f}원" for v in pnl['매출액']],[f"{v:.1f}%" for v in pnl['영업이익률']]))))
        fig.add_trace(go.Bar(x=pnl['영업이익'],y=pnl[group_col],name='영업이익',orientation='h',marker_color='#E8853D',opacity=0.8,text=[fmt_krw_short(v) for v in pnl['영업이익']],textposition='outside',textfont=dict(size=10),hovertemplate='%{y}<br>영업이익: %{customdata[0]}<br>영업이익률: %{customdata[1]}<extra></extra>',customdata=list(zip([f"{v:,.0f}원" for v in pnl['영업이익']],[f"{v:.1f}%" for v in pnl['영업이익률']]))))
        pnl_tvals,pnl_ttexts=krw_tickvals(pnl['매출액'])
        fig.update_layout(height=max(420,len(pnl)*50+140),barmode='group',margin=dict(l=180,r=100,t=30,b=40),legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=11)),xaxis=dict(title='금액',tickvals=pnl_tvals,ticktext=pnl_ttexts,tickfont=dict(size=11)),yaxis=dict(title='',tickfont=dict(size=10)))
        st.plotly_chart(fig, use_container_width=True, key=f"{tab_key}_chart")
        tbl_search=st.text_input(f"🔍 {group_col} 검색",key=f"{tab_key}_search")
        tbl=pnl.sort_values('매출액',ascending=False).reset_index(drop=True)
        if tbl_search: tbl=tbl[tbl[group_col].str.contains(tbl_search,case=False,na=False)]
        st.dataframe(tbl.style.format({'매출액':'{:,.0f}원','매출총이익':'{:,.0f}원','매출총이익률':'{:.1f}%','영업이익':'{:,.0f}원','영업이익률':'{:.1f}%'}),use_container_width=True,height=400)
    sub1,sub2,sub3=st.tabs(["대분류","중분류","소분류"])
    with sub1: _render_pnl(bw,'제품계층구조(대)',f"{pfx}_large")
    with sub2:
        sel=st.selectbox("대분류 선택",["전체"]+sorted(bw['제품계층구조(대)'].unique().tolist()),key=f"{pfx}_mid_sel")
        _render_pnl(bw if sel=="전체" else bw[bw['제품계층구조(대)']==sel],'제품계층구조(중)',f"{pfx}_medium")
    with sub3:
        c1,c2=st.columns(2)
        with c1: sel_l=st.selectbox("대분류",["전체"]+sorted(bw['제품계층구조(대)'].unique().tolist()),key=f"{pfx}_small_l")
        bw_s=bw if sel_l=="전체" else bw[bw['제품계층구조(대)']==sel_l]
        with c2: sel_m=st.selectbox("중분류",["전체"]+sorted(bw_s['제품계층구조(중)'].unique().tolist()),key=f"{pfx}_small_m")
        _render_pnl(bw_s if sel_m=="전체" else bw_s[bw_s['제품계층구조(중)']==sel_m],'제품계층구조(소)',f"{pfx}_small")

def render_material_pnl_table(kp=""):
    st.markdown("#### 자재별 손익 현황")
    if len(bw_filtered)==0: st.warning("BW 데이터 없음"); return
    bw=bw_filtered.copy()
    mat_pnl=bw.groupby(['자재','자재명']).agg(매출액=('I.매출액(FI기준)','sum'),매출원가=('II.매출원가','sum'),매출총이익=('III.매출총이익','sum'),판관비=('IV.판매비 및 관리비','sum'),영업이익=('V.영업이익I','sum'),판매수량=('판매수량','sum')).reset_index()
    mat_pnl['매출총이익률']=np.where(mat_pnl['매출액']!=0,mat_pnl['매출총이익']/mat_pnl['매출액']*100,0)
    mat_pnl['영업이익률']=np.where(mat_pnl['매출액']!=0,mat_pnl['영업이익']/mat_pnl['매출액']*100,0)
    mat_pnl=mat_pnl.sort_values('매출액',ascending=False).reset_index(drop=True)
    ms=st.text_input("🔍 자재명/코드 검색",key=f"{kp}_mat_search" if kp else "mat_search_main")
    if ms: mat_pnl=mat_pnl[mat_pnl.apply(lambda r:ms.lower() in str(r['자재명']).lower() or ms in str(r['자재']),axis=1)]
    def highlight_negative(val):
        if isinstance(val,(int,float)) and val<0: return 'color: #E74C3C; font-weight: 600'
        return ''
    st.dataframe(mat_pnl.style.format({'매출액':'{:,.0f}원','매출원가':'{:,.0f}원','매출총이익':'{:,.0f}원','매출총이익률':'{:.1f}%','판관비':'{:,.0f}원','영업이익':'{:,.0f}원','영업이익률':'{:.1f}%','판매수량':'{:,.0f}'}).map(highlight_negative,subset=['영업이익','영업이익률']),use_container_width=True,height=550)
# ============================================================
# 38개 차트 레지스트리
# ============================================================
CHART_REGISTRY = {
    "C01":{"name":"📈 월별 매출·주문건수 추이","tab":"종합 현황","fn":render_monthly_sales_trend},
    "C02":{"name":"🥧 회원구분별 매출 비중","tab":"종합 현황","fn":render_member_type_sales_donut},
    "C03":{"name":"🗺️ 지역별 매출","tab":"종합 현황","fn":render_region_sales_bar},
    "C04":{"name":"📅 일별 매출 추이","tab":"종합 현황","fn":render_daily_sales_trend},
    "C05":{"name":"📊 회원구분별×월별 매출 추이","tab":"매출 분석","fn":render_type_monthly_sales},
    "C06":{"name":"🏅 회원등급별 매출","tab":"매출 분석","fn":render_grade_sales_bar},
    "C07":{"name":"🔥 요일·시간대별 히트맵","tab":"매출 분석","fn":render_heatmap_dow_hour},
    "C08":{"name":"🏢 기관별 매출 현황 테이블","tab":"매출 분석","fn":render_org_sales_table},
    "C09":{"name":"📦 상품별 매출 TOP 20 파레토","tab":"상품 분석","fn":render_product_pareto},
    "C10":{"name":"📋 전체 상품 매출 현황 테이블","tab":"상품 분석","fn":render_product_sales_table},
    "C11":{"name":"🔀 회원구분별×상품 매출 크로스","tab":"상품 분석","fn":render_product_type_cross},
    "C12":{"name":"📉 월별 상품 매출 추이","tab":"상품 분석","fn":render_product_monthly_trend},
    "C13":{"name":"👥 월별 신규가입자 추이","tab":"회원 분석","fn":render_new_member_trend},
    "C14":{"name":"🎖️ 회원등급별 가입자 분포","tab":"회원 분석","fn":render_grade_member_bar},
    "C15":{"name":"⏱️ 가입 후 첫 주문까지 소요일","tab":"회원 분석","fn":render_first_order_days},
    "C16":{"name":"🔢 주문횟수 구간별 회원 분포","tab":"회원 분석","fn":render_order_count_dist},
    "C17":{"name":"🧩 코호트 리텐션 히트맵","tab":"회원 분석","fn":render_cohort_heatmap},
    "C18":{"name":"😴 휴면·활성 회원 분석","tab":"회원 분석","fn":render_dormant_analysis},
    "C19":{"name":"🔗 추천인 유형별 피추천인 수","tab":"추천인 분석","fn":render_referral_count_bar},
    "C20":{"name":"💰 추천인 유형별 피추천인 매출","tab":"추천인 분석","fn":render_referral_sales_donut},
    "C21":{"name":"📃 추천인별 현황 테이블","tab":"추천인 분석","fn":render_referral_table},
    "C22":{"name":"🤝 케어포 등급별 매출","tab":"케어포 멤버십","fn":render_carefor_grade_sales},
    "C23":{"name":"📬 케어포 등급별 주문건수","tab":"케어포 멤버십","fn":render_carefor_grade_orders},
    "C24":{"name":"📆 케어포 월별 매출 추이","tab":"케어포 멤버십","fn":render_carefor_monthly_trend},
    "C25":{"name":"🆕 케어포 등급별 신규가입 추이","tab":"케어포 멤버십","fn":render_carefor_new_member_trend},
    "C26":{"name":"💹 손익 워터폴","tab":"손익 분석","fn":render_pnl_waterfall},
    "C27":{"name":"📈 월별 손익 추이","tab":"손익 분석","fn":render_pnl_monthly_trend},
    "C28":{"name":"🏦 채널별 손익 비교","tab":"손익 분석","fn":render_channel_pnl},
    "C29":{"name":"📑 채널별 손익 상세 테이블","tab":"손익 분석","fn":render_channel_pnl_table},
    "C30":{"name":"🧾 판관비 항목별 비중","tab":"손익 분석","fn":render_sga_composition},
    "C31":{"name":"📊 월별 판관비 구성 추이","tab":"손익 분석","fn":render_sga_monthly_trend},
    "C32":{"name":"🏷️ 제품계층구조별 수익성 분석","tab":"손익 분석","fn":render_product_pnl_hierarchy},
    "C33":{"name":"🔩 자재별 손익 현황 테이블","tab":"손익 분석","fn":render_material_pnl_table},
    "C34":{"name":"🏪 대리점 피추천인 매출 및 판매수수료 집계","tab":"추천인 분석","fn":render_dealer_commission},
}

# ============================================================
# 탭 구성
# ============================================================
tab1,tab2,tab3,tab4,tab5,tab6,tab7,tab8,tab9 = st.tabs([
    "📋 종합 현황","💵 매출 분석","📦 상품 분석","👥 회원 분석",
    "🔗 추천인 분석","🤝 케어포 멤버십","📈 손익 분석","🏥 일차의료 시범기관","⚙️ 커스텀 뷰"
])

# ============================================================
# Tab 1. 종합 현황
# ============================================================
with tab1:
    ts=filtered['판매합계금액'].sum(); to_=filtered['주문 ID'].nunique(); tb=filtered['주문자 ID'].nunique()
    tm=len(members); nm=len(filtered_members); ao=ts/to_ if to_>0 else 0
    cols=st.columns(6)
    for col,(l,v,u) in zip(cols,[("총 매출액",fmt_krw_short(ts),"원"),("총 주문건수",fmt_num(to_),"건"),("총 회원수",fmt_num(tm),"처"),("주문회원수",fmt_num(tb),"처"),("신규 가입회원",fmt_num(nm),"처"),("객단가",fmt_krw_short(ao),"원")]):
        col.markdown(kpi_card(l,v,u),unsafe_allow_html=True)
    render_monthly_sales_trend()
    cl,cr=st.columns(2)
    with cl: render_member_type_sales_donut()
    with cr: render_region_sales_bar()
    render_daily_sales_trend()

# ============================================================
# Tab 2. 매출 분석
# ============================================================
with tab2:
    render_type_monthly_sales()
    render_grade_sales_bar()
    render_heatmap_dow_hour()
    render_org_sales_table()

# ============================================================
# Tab 3. 상품 분석
# ============================================================
with tab3:
    render_product_pareto()
    render_product_sales_table()
    render_product_type_cross()
    render_product_monthly_trend()

# ============================================================
# Tab 4. 회원 분석
# ============================================================
with tab4:
    mo_df=orders.groupby('주문자 ID').agg(첫주문일=('주문일','min'),주문건수=('주문 ID','nunique'),주문월수=('주문월','nunique')).reset_index()
    conv=members[members['아이디'].isin(orders['주문자 ID'].unique())]; conv_r=len(conv)/len(members)*100 if len(members)>0 else 0
    rep=mo_df[mo_df['주문건수']>=2]; rep_r=len(rep)/len(mo_df)*100 if len(mo_df)>0 else 0
    r3m=orders[orders['주문일']>=orders['주문일'].max()-pd.DateOffset(months=3)]; act=r3m['주문자 ID'].nunique()
    cols=st.columns(5)
    for col,(l,v,u) in zip(cols,[("총 회원수",fmt_num(len(members)),"처"),("신규 가입회원",fmt_num(len(filtered_members)),"처"),("구매전환율",fmt_pct(conv_r),""),("재구매율",fmt_pct(rep_r),""),("활성회원(3개월)",fmt_num(act),"처")]):
        col.markdown(kpi_card(l,v,u),unsafe_allow_html=True)
    st.markdown("#### 🔍 회원 상세 검색")
    mem_search=st.text_input("상호명, 아이디, 담당자명으로 검색",key="member_detail_search",placeholder="예: 대상병원, 대상요양원 등")
    if mem_search:
        mask=members.apply(lambda r:mem_search.lower() in str(r.get('상호명','')).lower() or mem_search.lower() in str(r.get('아이디','')).lower() or mem_search.lower() in str(r.get('담당자 이름','')).lower(),axis=1)
        search_members=members[mask].copy()
        if len(search_members)==0: st.warning("검색 결과가 없습니다.")
        else:
            order_agg=orders.groupby('주문자 ID').agg(총매출=('판매합계금액','sum'),주문건수=('주문 ID','nunique'),첫주문일=('주문일자','min'),최근주문일=('주문일자','max')).reset_index()
            order_agg['객단가']=(order_agg['총매출']/order_agg['주문건수']).round(0)
            top_products=orders.groupby(['주문자 ID','상품명'])['판매합계금액'].sum().reset_index().sort_values(['주문자 ID','판매합계금액'],ascending=[True,False])
            top_products=top_products.groupby('주문자 ID').head(3).groupby('주문자 ID')['상품명'].apply(lambda x:' / '.join(x)).reset_index(); top_products.columns=['주문자 ID','주요 구매상품']
            ref_info=referrals_df.groupby('피추천인 사업자 번호').first()[['추천인','회원그룹']].reset_index(); ref_info.columns=['사업자번호','추천인명','추천인유형']
            result=search_members[['아이디','상호명','사업자번호','회원타입','회원등급','가입일','담당자 이름','휴대폰','주소']].copy()
            result['가입일']=result['가입일'].dt.strftime('%Y-%m-%d')
            result=result.merge(order_agg,left_on='아이디',right_on='주문자 ID',how='left').drop(columns=['주문자 ID'],errors='ignore')
            result=result.merge(top_products,left_on='아이디',right_on='주문자 ID',how='left').drop(columns=['주문자 ID'],errors='ignore')
            result=result.merge(ref_info,on='사업자번호',how='left')
            for c in ['총매출','객단가']: result[c]=result[c].fillna(0)
            result['주문건수']=result['주문건수'].fillna(0).astype(int)
            for c in ['주요 구매상품','추천인명','추천인유형']: result[c]=result[c].fillna('-')
            display_cols=['아이디','상호명','담당자 이름','회원타입','회원등급','가입일','총매출','주문건수','객단가','첫주문일','최근주문일','주요 구매상품','추천인명','추천인유형','휴대폰','주소']
            result=result[[c for c in display_cols if c in result.columns]].sort_values('총매출',ascending=False).reset_index(drop=True)
            st.markdown(f"**검색 결과: {len(result)}건**")
            st.dataframe(result.style.format({'총매출':'{:,.0f}원','주문건수':'{:,.0f}건','객단가':'{:,.0f}원'}),use_container_width=True,height=400)
    render_new_member_trend()
    render_grade_member_bar()
    cl,cr=st.columns(2)
    with cl: render_first_order_days()
    with cr: render_order_count_dist()
    render_cohort_heatmap()
    render_dormant_analysis()
    # 휴면 회원 목록 테이블 (Tab4 전용)
    base_date=orders['주문일'].max()
    last_order_t4=orders.groupby('주문자 ID')['주문일'].max().reset_index(); last_order_t4.columns=['아이디','마지막주문일']
    dormant_df=members.copy(); dormant_df=dormant_df.merge(last_order_t4,on='아이디',how='left')
    dormant_df['휴면경과일']=(base_date-dormant_df['마지막주문일']).dt.days
    def classify_status_t4(row):
        if pd.isna(row['마지막주문일']): return '미구매'
        d=row['휴면경과일']
        if d<=90: return '활성'
        elif d<=180: return '단기휴면'
        elif d<=365: return '중기휴면'
        else: return '장기휴면'
    dormant_df['활성구분']=dormant_df.apply(classify_status_t4,axis=1)
    dormant_df['마지막주문일']=dormant_df['마지막주문일'].dt.strftime('%Y-%m-%d')
    dormant_df['가입일']=dormant_df['가입일'].dt.strftime('%Y-%m-%d')
    status_order_t4=['활성','단기휴면','중기휴면','장기휴면','미구매']
    st.markdown("##### 회원 목록")
    col_filter,col_type,col_grade=st.columns(3)
    with col_filter: status_filter=st.multiselect("활성 구분 필터",status_order_t4,default=[],placeholder="전체",key="dormant_status_filter")
    with col_type:
        dormant_type_opts=sorted(dormant_df['회원타입'].dropna().unique().tolist())
        dormant_type=st.multiselect("회원타입",dormant_type_opts,default=[],placeholder="전체",key="dormant_type_filter")
    with col_grade:
        dormant_grade_opts=sorted(dormant_df['회원등급'].dropna().unique().tolist())
        dormant_grade=st.multiselect("회원등급",dormant_grade_opts,default=[],placeholder="전체",key="dormant_grade_filter")
    display_cols_d=['추천인','상호명','아이디','휴대폰','주소','가입일','회원타입','회원등급','마지막주문일','휴면경과일','SMS 수신동의','활성구분']
    exist_cols_d=[c for c in display_cols_d if c in dormant_df.columns]
    tbl=dormant_df[exist_cols_d].copy()
    if status_filter: tbl=tbl[tbl['활성구분'].isin(status_filter)]
    if dormant_type: tbl=tbl[tbl['회원타입'].isin(dormant_type)]
    if dormant_grade: tbl=tbl[tbl['회원등급'].isin(dormant_grade)]
    tbl=tbl.sort_values('휴면경과일',ascending=False,na_position='last').reset_index(drop=True)
    st.markdown(f"**{len(tbl):,}건**")
    def to_excel(df):
        output=io.BytesIO()
        with pd.ExcelWriter(output,engine='openpyxl') as writer: df.to_excel(writer,index=False,sheet_name='휴면회원분석')
        return output.getvalue()
    st.download_button(label="📥 엑셀 다운로드",data=to_excel(tbl),file_name=f"휴면회원분석_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.dataframe(tbl,use_container_width=True,height=500)

# ============================================================
# Tab 5. 추천인 분석
# ============================================================
with tab5:
    rdf=_build_referral_data()
    cols=st.columns(3)
    for col,(l,v,u) in zip(cols,[("총 추천인 수",fmt_num(len(rdf)),"회원"),("총 피추천인 수",fmt_num(rdf['피추천인수'].sum()),"회원"),("추천인당 평균 피추천인",f"{rdf['피추천인수'].mean():.1f}" if len(rdf)>0 else "0","회원")]): col.markdown(kpi_card(l,v,u),unsafe_allow_html=True)
    cl,cr=st.columns(2)
    with cl: render_referral_count_bar()
    with cr: render_referral_sales_donut()
    st.markdown("#### 추천인별 현황")
    rtf=st.selectbox("추천인 유형 필터",["전체","영업팀","대리점","케어포"],key="ref_type"); dr=rdf.copy()
    if rtf!="전체": dr=dr[dr['유형']==rtf]
    dr=dr.sort_values('피추천인매출',ascending=False).reset_index(drop=True)
    sr=st.text_input("🔍 추천인 검색",key="ref_search")
    if sr: dr=dr[dr.apply(lambda r:sr.lower() in str(r).lower(),axis=1)]
    st.dataframe(dr.style.format({'피추천인수':'{:,.0f}','피추천인매출':'{:,.0f}원'}),use_container_width=True,height=550)
    # ── 대리점 판매수수료 집계 ──
    st.markdown("---")
    render_dealer_commission()
    st.markdown("---")
    render_dealer_commission_forecast()
    
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
    c1,c2=st.columns(2)
    with c1: render_carefor_grade_sales()
    with c2: render_carefor_grade_orders()
    render_carefor_monthly_trend()
    render_carefor_new_member_trend()

# ============================================================
# Tab 7. 손익 분석 (BW)
# ============================================================
with tab7:
    if len(bw_data)==0:
        st.warning("⚠️ BW 손익 데이터가 없습니다. 엑셀 파일에 'BW' 시트를 추가해주세요.")
    else:
        bw=bw_filtered.copy()
        bw_channels=sorted(bw['채널'].unique().tolist())
        sel_channels=st.multiselect("채널 필터",bw_channels,default=[],placeholder="전체",key="bw_channel")
        if sel_channels: bw=bw[bw['채널'].isin(sel_channels)]
        prod_large=sorted(bw['제품계층구조(대)'].dropna().unique().tolist())
        sel_prod_l=st.multiselect("제품계층구조(대) 필터",prod_large,default=[],placeholder="전체",key="bw_prod_l")
        if sel_prod_l: bw=bw[bw['제품계층구조(대)'].isin(sel_prod_l)]
        rev=bw['I.매출액(FI기준)'].sum(); cogs=bw['II.매출원가'].sum(); gp=bw['III.매출총이익'].sum(); sga=bw['IV.판매비 및 관리비'].sum(); oi=bw['V.영업이익I'].sum(); oi_rate=(oi/rev*100) if rev!=0 else 0
        cols=st.columns(6)
        for col,(l,v,u) in zip(cols,[("매출액",fmt_krw_short(rev),"원"),("매출원가",fmt_krw_short(cogs),"원"),("매출총이익",fmt_krw_short(gp),"원"),("판관비",fmt_krw_short(sga),"원"),("영업이익",fmt_krw_short(oi),"원"),("영업이익률",fmt_pct(oi_rate),"")]):
            col.markdown(kpi_card(l,v,u),unsafe_allow_html=True)
        render_pnl_waterfall()
        render_pnl_monthly_trend()
        render_channel_pnl()
        render_channel_pnl_table()
        c1,c2=st.columns(2)
        with c1: render_sga_composition()
        with c2: render_sga_monthly_trend()
        render_product_pnl_hierarchy()
        render_material_pnl_table()

# ============================================================
# Tab 8. 일차의료 시범기관
# ============================================================
with tab8:
    pilot_df=load_pilot_clinics()
    if pilot_df.empty:
        st.warning("⚠️ 일차의료 시범기관 데이터를 불러올 수 없습니다.")
    else:
        type_opts_p=sorted(pilot_df['사업유형'].unique().tolist())
        sel_pilot_type=st.multiselect("사업유형 필터",type_opts_p,default=type_opts_p,key="pilot_type")
        pf=pilot_df[pilot_df['사업유형'].isin(sel_pilot_type)] if sel_pilot_type else pilot_df
        match_df=match_pilot_clinics(pf,members,orders)
        total_clinics=len(pf); matched_count=len(match_df)
        confirmed=len(match_df[match_df['매칭등급']=='확정']) if not match_df.empty else 0
        candidate=len(match_df[match_df['매칭등급']=='후보']) if not match_df.empty else 0
        matched_revenue=match_df.drop_duplicates(subset='아이디')['총매출'].sum() if not match_df.empty else 0
        match_rate=(matched_count/total_clinics*100) if total_clinics>0 else 0
        cols=st.columns(6)
        for col,(l,v,u) in zip(cols,[("시범기관 수",fmt_num(total_clinics),"곳"),("B2B몰 가입",fmt_num(matched_count),"곳"),("가입률",fmt_pct(match_rate),""),("확정 매칭",fmt_num(confirmed),"곳"),("후보 매칭",fmt_num(candidate),"곳"),("가입기관 매출",fmt_krw_short(matched_revenue),"원")]):
            col.markdown(kpi_card(l,v,u),unsafe_allow_html=True)
        c1,c2=st.columns(2)
        with c1:
            st.markdown("#### 사업유형별 기관 수")
            type_summary=pf['사업유형'].value_counts().reset_index(); type_summary.columns=['사업유형','기관수']
            if not match_df.empty:
                type_matched=match_df.groupby('사업유형').size().reset_index(name='가입수')
                type_summary=type_summary.merge(type_matched,on='사업유형',how='left')
            else: type_summary['가입수']=0
            type_summary['가입수']=type_summary['가입수'].fillna(0).astype(int); type_summary['미가입']=type_summary['기관수']-type_summary['가입수']
            fig=go.Figure()
            fig.add_trace(go.Bar(x=type_summary['사업유형'],y=type_summary['가입수'],name='B2B 가입',marker_color='#27AE60',text=[fmt_num(v) for v in type_summary['가입수']],textposition='inside',textfont=dict(size=11),hovertemplate='%{x}<br>가입: %{y}곳<extra></extra>'))
            fig.add_trace(go.Bar(x=type_summary['사업유형'],y=type_summary['미가입'],name='미가입',marker_color='#BDC3C7',text=[fmt_num(v) for v in type_summary['미가입']],textposition='inside',textfont=dict(size=11),hovertemplate='%{x}<br>미가입: %{y}곳<extra></extra>'))
            fig.update_layout(height=450,barmode='stack',margin=dict(l=60,r=30,t=30,b=40),legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=11)))
            st.plotly_chart(fig,use_container_width=True,key="pilot_type_bar")
        with c2:
            st.markdown("#### 매칭 현황")
            has_revenue=match_df[match_df['총매출']>0]['아이디'].nunique() if not match_df.empty else 0
            no_revenue=match_df[match_df['총매출']==0]['아이디'].nunique() if not match_df.empty else 0
            status_df=pd.DataFrame({'구분':['매출 발생','매출 미발생(휴면)','미가입'],'수':[has_revenue,no_revenue,total_clinics-matched_count]})
            status_df=status_df[status_df['수']>0]
            fig=make_donut(status_df,'구분','수',colors=['#27AE60','#E8853D','#BDC3C7'],value_label='기관수',unit='곳')
            fig.update_layout(height=450); fig.layout.annotations[0].text=f"<b>매칭기관</b><br>{fmt_num(matched_count)}곳"
            st.plotly_chart(fig,use_container_width=True,key="pilot_match_donut")
        st.markdown("#### 지역별 시범기관 분포")
        region_all=pf['시도'].value_counts().reset_index(); region_all.columns=['지역','시범기관수']
        if not match_df.empty:
            region_matched=match_df.copy(); region_matched['지역']=region_matched['주소_공공'].apply(normalize_sido)
            region_matched=region_matched.groupby('지역').size().reset_index(name='가입기관수')
            region_all=region_all.merge(region_matched,on='지역',how='left')
        else: region_all['가입기관수']=0
        region_all['가입기관수']=region_all['가입기관수'].fillna(0).astype(int); region_all['미가입']=region_all['시범기관수']-region_all['가입기관수']; region_all=region_all.sort_values('시범기관수')
        fig=go.Figure()
        fig.add_trace(go.Bar(y=region_all['지역'],x=region_all['가입기관수'],name='B2B 가입',orientation='h',marker_color='#27AE60',hovertemplate='%{y}<br>가입: %{x}곳<extra></extra>'))
        fig.add_trace(go.Bar(y=region_all['지역'],x=region_all['미가입'],name='미가입',orientation='h',marker_color='#BDC3C7',hovertemplate='%{y}<br>미가입: %{x}곳<extra></extra>'))
        fig.update_layout(height=max(450,len(region_all)*28+140),barmode='stack',margin=dict(l=100,r=30,t=30,b=40),legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=11)))
        st.plotly_chart(fig,use_container_width=True,key="pilot_region_bar")
        if not match_df.empty and matched_revenue>0:
            st.markdown("#### 매칭 기관 매출 TOP 20")
            top_dedup=match_df[match_df['총매출']>0].copy()
            top_dedup['사업유형목록']=top_dedup.groupby('아이디')['사업유형'].transform(lambda x:'/'.join(sorted(x.unique())))
            top_dedup=top_dedup.drop_duplicates(subset='아이디').sort_values('총매출',ascending=False).head(20)
            if len(top_dedup)>0:
                top_dedup=top_dedup.sort_values('총매출',ascending=True)
                type_colors={'만성질환관리':'#3366CC','방문진료':'#E8853D','한의방문진료':'#27AE60','만성질환관리/방문진료':'#9B59B6','만성질환관리/한의방문진료':'#1ABC9C'}
                fig=go.Figure()
                fig.add_trace(go.Bar(x=top_dedup['총매출'].values,y=top_dedup['상호명_B2B'].values,orientation='h',marker_color=[type_colors.get(t,'#3366CC') for t in top_dedup['사업유형목록']],showlegend=False,text=[fmt_krw_short(v) for v in top_dedup['총매출']],textposition='outside',textfont=dict(size=10),hovertemplate='%{y}<br>매출: %{customdata[0]}<br>사업유형: %{customdata[1]}<extra></extra>',customdata=list(zip([f"{v:,.0f}원" for v in top_dedup['총매출']],top_dedup['사업유형목록']))))
                for label,color in [('만성질환관리','#3366CC'),('방문진료','#E8853D'),('한의방문진료','#27AE60'),('만성질환관리/방문진료','#9B59B6'),('만성질환관리/한의방문진료','#1ABC9C')]: fig.add_trace(go.Bar(x=[None],y=[None],marker_color=color,name=label,showlegend=True))
                top_tvals,top_ttexts=krw_tickvals(top_dedup['총매출'])
                fig.update_layout(height=max(450,len(top_dedup)*28+140),margin=dict(l=180,r=80,t=30,b=40),showlegend=True,legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0,font=dict(size=11)),xaxis=dict(title='매출액',tickvals=top_tvals,ticktext=top_ttexts,tickfont=dict(size=11)))
                st.plotly_chart(fig,use_container_width=True,key="pilot_top20")
        st.markdown("#### 매칭 기관 상세")
        match_filter=st.selectbox("매칭등급 필터",["전체","확정","후보"],key="pilot_match_filter")
        display_df=match_df.copy() if not match_df.empty else pd.DataFrame()
        if not display_df.empty:
            if match_filter!="전체": display_df=display_df[display_df['매칭등급']==match_filter]
            display_df=display_df.sort_values('총매출',ascending=False).reset_index(drop=True)
            search_pilot=st.text_input("🔍 기관명/상호명 검색",key="pilot_search")
            if search_pilot: display_df=display_df[display_df.apply(lambda r:search_pilot.lower() in str(r['기관명']).lower() or search_pilot.lower() in str(r['상호명_B2B']).lower(),axis=1)]
            st.dataframe(display_df.style.format({'총매출':'{:,.0f}원','주문건수':'{:,.0f}건'}),use_container_width=True,height=550)
        else: st.info("매칭된 기관이 없습니다.")
        st.markdown("#### 미매칭 시범기관 목록 (잠재 영업 대상)")
        if not match_df.empty:
            matched_keys=set(zip(match_df['기관명'],match_df['사업유형']))
            unmatched=pf[~pf.apply(lambda r:(r['기관명'],r['사업유형']) in matched_keys,axis=1)][['사업유형','기관명','기관구분','주소','전화번호']].reset_index(drop=True)
        else: unmatched=pf[['사업유형','기관명','기관구분','주소','전화번호']].reset_index(drop=True)
        st.markdown(f"**미매칭: {len(unmatched)}곳**")
        search_unmatched=st.text_input("🔍 미매칭 기관 검색",key="pilot_unmatched_search")
        if search_unmatched: unmatched=unmatched[unmatched.apply(lambda r:search_unmatched.lower() in str(r).lower(),axis=1)]
        st.dataframe(unmatched,use_container_width=True,height=450)

# ============================================================
# Tab 9. 커스터마이징
# ============================================================
with tab9:
    from streamlit_sortables import sort_items
    st.markdown("### ⚙️ 대시보드 커스터마이징")
    st.markdown("차트를 선택하고 **▶ 순서 조정으로 이동** 버튼을 누르세요.")

    if 'custom_step' not in st.session_state: st.session_state.custom_step='select'
    if 'custom_selected' not in st.session_state: st.session_state.custom_selected=[]
    if 'custom_order' not in st.session_state: st.session_state.custom_order=[]

    tab_groups={}
    for cid,info in CHART_REGISTRY.items(): tab_groups.setdefault(info["tab"],[]).append((cid,info["name"]))

    # ── Step 1: 선택 (form으로 rerun 억제) ──
    if st.session_state.custom_step in ('select','sort','view'):
        with st.form("chart_select_form"):
            st.markdown("#### 1단계: 차트 선택")
            new_selected=[]
            tab_cols=st.columns(2)
            tab_list=list(tab_groups.items())
            for i,(tab_name,charts) in enumerate(tab_list):
                with tab_cols[i%2]:
                    st.markdown(f"**{tab_name}**")
                    for cid,cname in charts:
                        checked=st.checkbox(cname,value=(cid in st.session_state.custom_selected),key=f"chk_{cid}")
                        if checked: new_selected.append(cid)
            submitted=st.form_submit_button("▶ 순서 조정으로 이동",type="primary",use_container_width=True)
            if submitted:
                st.session_state.custom_selected=new_selected
                st.session_state.custom_step='sort'; st.rerun()

    # ── Step 2: 순서 조정 ──
    if st.session_state.custom_step in ('sort','view'):
        st.markdown("---")
        st.markdown("#### 2단계: 순서 조정 (드래그앤드롭)")
        if not st.session_state.custom_selected:
            st.warning("선택된 차트가 없습니다.")
        else:
            labels=[CHART_REGISTRY[cid]["name"] for cid in st.session_state.custom_selected]
            sorted_result=sort_items(labels,direction="vertical",key="custom_sort")
            col_back,col_run=st.columns([1,3])
            with col_back:
                if st.button("← 다시 선택",use_container_width=True):
                    st.session_state.custom_step='select'; st.rerun()
            with col_run:
                if st.button("▶ 대시보드 보기",type="primary",use_container_width=True):
                    st.session_state.custom_order=sorted_result
                    st.session_state.custom_step='view'; st.rerun()

    # ── Step 3: 렌더링 ──
    if st.session_state.custom_step=='view':
        st.markdown("---")
        label_to_cid={info["name"]:cid for cid,info in CHART_REGISTRY.items()}
        col_back2,_=st.columns([1,3])
        with col_back2:
            if st.button("← 순서 조정으로 돌아가기",use_container_width=True):
                st.session_state.custom_step='sort'; st.rerun()
        st.markdown(f"## 📊 커스텀 대시보드  <span style='font-size:0.9rem;color:#94a3b8;'>({len(st.session_state.custom_order)}개 차트)</span>",unsafe_allow_html=True)
        for idx,label in enumerate(st.session_state.custom_order):
            cid=label_to_cid.get(label)
            if cid and cid in CHART_REGISTRY:
                kp=f"custom_{cid}_{idx}"  # 고유 key prefix → 각 render 함수 내부에서 직접 사용
                with st.container():
                    CHART_REGISTRY[cid]["fn"](kp=kp)
                st.markdown("---")

# ============================================================
# 푸터
# ============================================================
st.markdown("---")
st.markdown(f"<p style='text-align:center;color:#94a3b8;font-size:0.85rem;'>© 대상웰라이프 B2B몰 대시보드 · 데이터 기준: {pd.Timestamp.now().strftime('%Y년 %m월 %d일')}</p>",unsafe_allow_html=True)
