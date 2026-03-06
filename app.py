import streamlit as st
import pandas as pd
import plotly.express as px

# ==========================================
# 0. 페이지 설정 (가장 먼저 호출)
# ==========================================
st.set_page_config(
    page_title="웰스FA VVIP 수수료 대시보드",
    page_icon="💰",
    layout="wide"
)

# ==========================================
# 1. VVIP 다크 테마 CSS 주입
# ==========================================
CUSTOM_CSS = """
<style>
    :root {
        --bg-main: #0A1128;
        --bg-card: #16203B;
        --color-gold: #D4AF37;
        --color-text: #E2E8F0;
        --color-danger: #E63946;
    }
    .stApp, .main, [data-testid="stAppViewContainer"] {
        background-color: var(--bg-main) !important;
        color: var(--color-text) !important;
    }
    /* 사이드바 완전 숨기기 */
    [data-testid="stSidebar"], [data-testid="collapsedControl"] {
        display: none !important;
    }
    .block-container { padding-left: 2rem !important; padding-right: 2rem !important; }
    p, label, .stMarkdown > div, .stText {
        color: var(--color-text);
        text-shadow: none !important;
    }
    /* ★ 메인 타이틀(h1) 메탈릭 골드 강제주입 - Streamlit 기본 흰색 차단 */
    .main h1, [data-testid="stMarkdownContainer"] h1 {
        background: linear-gradient(90deg, #BF953F, #FCF6BA, #D4AF37, #FBF5B7, #AA771C) !important;
        -webkit-background-clip: text !important;
        -webkit-text-fill-color: transparent !important;
        color: transparent !important;
        filter: drop-shadow(0px 4px 10px rgba(212,175,55,0.4)) !important;
        text-shadow: none !important;
        font-weight: 900 !important;
    }
    /* ★ 단위 '원' 축소 클래스 */
    .unit-won {
        font-size: 14px !important;
        color: #94A3B8 !important;
        font-weight: 400 !important;
    }
    /* KPI 카드 (업적 섹션) */
    .kpi-card {
        background: linear-gradient(180deg, var(--bg-card) 0%, #0A1128 100%) !important;
        border-radius: 12px !important;
        padding: 36px 28px 28px 28px !important;
        border: 1px solid rgba(255, 255, 255, 0.05) !important;
        border-top: 4px solid var(--color-gold) !important;
        overflow: hidden !important;
        text-align: center !important;
        margin-bottom: 20px;
    }
    .kpi-title { color: #A0AEC0 !important; font-weight:700; font-size:1.4rem; margin-bottom:5px; }
    .kpi-value { color: var(--color-gold) !important; font-weight:900; font-size:3.2rem; margin:0; }

    /* 수수료 대시보드 존 카드 (더블 컨테이너) */
    .kpi-card-zone {
        background: linear-gradient(180deg, var(--bg-card) 0%, #0A1128 100%) !important;
        border-radius: 12px !important;
        padding: 36px 28px 28px 28px !important;
        border: 1px solid rgba(255, 255, 255, 0.05) !important;
        border-top: 4px solid var(--color-gold) !important;
        overflow: hidden !important;
        text-align: center !important;
        margin-bottom: 15px;
    }
    .zone-title { color: #A0AEC0 !important; font-weight:700; font-size:1.4rem; margin-bottom:5px; }
    .zone-value { color: var(--color-gold) !important; font-weight:900; font-size:4.0rem; margin:0; }
    .zone-value-sub { color: var(--color-gold) !important; font-weight:900; font-size:2.8rem; margin:0; }

    /* 테이블 격자선 전면 철거 (VVIP 하이엔드) */
    table, th, td, [data-testid="stDataFrame"] > div {
        border-left: none !important;
        border-right: none !important;
        border-top: none !important;
    }
    th { border-bottom: 1px solid rgba(255, 255, 255, 0.1) !important; }
    td { border-bottom: 1px solid rgba(255, 255, 255, 0.05) !important; }
    table {
        width: 100%; border-collapse: collapse;
        background-color: var(--bg-card) !important;
        color: var(--color-text) !important;
    }
    th, td { background-color: var(--bg-card) !important; color: var(--color-text) !important; border: none !important; }
    th { border-bottom: 2px solid var(--color-gold) !important; padding: 12px 8px; text-align: center; font-weight: 700; }
    td { padding: 10px 8px; border-bottom: 1px solid rgba(255, 255, 255, 0.05) !important; }
    tr:hover td { background-color: rgba(255, 255, 255, 0.05) !important; }

    /* 상세보기 토글 버튼 */
    .table-toggle-btn {
        background-color: rgba(22, 32, 59, 0.8);
        border-radius: 8px; padding: 10px 16px;
        font-weight: 500; font-size: 1.0rem;
        color: var(--color-text); display: inline-block;
        margin-bottom: 12px; border: 1px solid rgba(255, 255, 255, 0.05);
        transition: color 0.2s ease, background-color 0.2s ease; cursor: pointer;
    }
    .table-toggle-btn:hover { color: var(--color-gold); background-color: rgba(255, 255, 255, 0.05); }
    .table-toggle-btn::after { content: ' ▼'; font-size: 0.85em; margin-left: 6px; opacity: 0.8; }
    .table-toggle-container { padding: 5px 0 10px 0 !important; border: none !important; position: relative; z-index: 10; text-align: left; }
    div.stMarkdown:has(.table-toggle-container) { margin-bottom: -1rem !important; position: relative; z-index: 10; }

    /* Streamlit 데이터프레임 배경 투명화 */
    [data-testid="stDataFrame"] > div {
        background-color: transparent !important; border: none !important; box-shadow: none !important; padding: 0px !important;
    }
    [data-testid="stDataFrame"] td { background-color: transparent !important; color: var(--color-text) !important; }
    [data-testid="stDataFrame"] th { background-color: transparent !important; color: var(--color-text) !important; border-bottom: 2px solid var(--color-gold) !important; }

    /* 비밀번호 입력 박스 커스텀 */
    .pw-box {
        max-width: 420px; margin: 10vh auto; text-align: center;
        background: var(--bg-card); border-radius: 20px;
        padding: 50px 40px; border-top: 4px solid var(--color-gold);
        box-shadow: 0 20px 60px rgba(0,0,0,0.7);
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ==========================================
# 2. 비밀번호 게이트 (통과 시에만 메인 출력)
# ==========================================
CORRECT_PASSWORD = "ksg8984"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("""
    <div class="pw-box">
        <h2 style="color:#D4AF37; font-size:2.2rem; font-weight:900; margin-bottom:8px;">💰 WealthFA</h2>
        <p style="color:#94A3B8; font-size:14px; margin-bottom:30px;">VVIP 수수료 분석 시스템</p>
    </div>
    """, unsafe_allow_html=True)

    pw = st.text_input("🔐 비밀번호를 입력하세요", type="password", placeholder="비밀번호 입력 후 Enter")
    if pw:
        if pw == CORRECT_PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("❌ 비밀번호가 올바르지 않습니다.")
    st.stop()


# ==========================================
# 3. 유틸 함수
# ==========================================
def disp_dataframe(df, *args, **kwargs):
    format_dict = {c: "{:,.0f}" for c in df.columns if pd.api.types.is_numeric_dtype(df[c])}
    def color_negative_red(val):
        if pd.api.types.is_number(val) and val < 0:
            return 'color: #E63946'
        return ''
    styled_df = df.style.format(format_dict, na_rep="").map(color_negative_red)
    return st.dataframe(styled_df, *args, **kwargs)


# ==========================================
# 4. 데이터 로드 (상대 경로 사용)
# ==========================================
DATA_PATH = "./data/최종병합_수수료명세서_압축판.xlsx"

@st.cache_data
def load_data():
    try:
        df_life   = pd.read_excel(DATA_PATH, sheet_name='생명보험사_202602')
        df_damage = pd.read_excel(DATA_PATH, sheet_name='손해보험사_202602')
    except FileNotFoundError:
        st.error(f"❌ 데이터 파일을 찾을 수 없습니다: `{DATA_PATH}`\n\n`data/` 폴더에 엑셀 파일을 넣어주세요.")
        st.stop()

    df_life['보험군']   = '생명보험'
    df_damage['보험군'] = '손해보험'
    df_all = pd.concat([df_life, df_damage], ignore_index=True)

    # 날짜 정제
    if '계약일자' in df_all.columns:
        df_all['계약일자_정제'] = pd.to_datetime(df_all['계약일자'], errors='coerce').dt.strftime('%Y-%m-%d')
    else:
        df_all['계약일자_정제'] = ''

    # 수치형 컬럼 처리
    numeric_cols = ['지사수수료', '업적지표1', '업적지표2', '업적지표3']
    for c in numeric_cols:
        if c in df_all.columns:
            df_all[c] = pd.to_numeric(df_all[c], errors='coerce').fillna(0)

    text_cols = ['증권번호', '계약일자', '계약자', 'FC명', '제휴사명', '상품명', '상품군', '지급구분']
    for c in text_cols:
        if c in df_all.columns:
            df_all[c] = df_all[c].astype(str)

    if '상품군' not in df_all.columns:
        df_all['상품군'] = '분류없음'
    df_all['상품군']  = df_all['상품군'].fillna('분류없음').astype(str).str.strip()
    df_all['FC명']    = df_all['FC명'].fillna('알수없음')
    df_all['제휴사명'] = df_all['제휴사명'].fillna('알수없음')
    return df_all


df = load_data()


# 사이드바 없이 전체 데이터 사용
filtered_df = df.copy()


# ==========================================
# 6. 메인 타이틀
# ==========================================
title_html = """
<div style="margin-bottom: 2rem;">
    <h1 style="
        font-size: 3.5rem; font-weight: 900;
        color: transparent !important;
        background: linear-gradient(90deg, #BF953F 0%, #FCF6BA 25%, #D4AF37 50%, #FBF5B7 75%, #AA771C 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent !important;
        filter: drop-shadow(0px 4px 10px rgba(212, 175, 55, 0.4));
        margin-bottom: 0.5rem; letter-spacing: -1px;
    ">웰스FA 26년 2월 수수료 정밀분석</h1>
    <p style="font-size: 14px; color: #94A3B8 !important; font-weight: 500; margin-top: 0; text-shadow: none !important;">
        ※ 본 수수료명세서는 1월 영업 업적을 바탕으로 산출된 데이터입니다.
    </p>
</div>
"""
st.markdown(title_html, unsafe_allow_html=True)
st.markdown("---")


# ==========================================
# 7. 사전 데이터 연산
# ==========================================
life_df   = filtered_df[filtered_df['보험군'] == '생명보험']
damage_df = filtered_df[filtered_df['보험군'] == '손해보험']

is_life_new = life_df['지급구분'].astype(str).str.contains('신계약', na=False)
is_dmg_new  = damage_df['지급구분'].astype(str).str.contains('신계약', na=False)
is_dmg_ret  = damage_df['지급구분'].astype(str).str.contains('유지', na=False)
is_dmg_etc  = damage_df['지급구분'].astype(str).str.contains('일반|자동차|환수', na=False, regex=True)

life_new_df = life_df[is_life_new]
life_ret_df = life_df[~is_life_new]
dmg_new_df  = damage_df[is_dmg_new]
dmg_ret_df  = damage_df[is_dmg_ret]
dmg_etc_df  = damage_df[is_dmg_etc]

val_life_new = life_new_df['지사수수료'].sum()
val_life_ret = life_ret_df['지사수수료'].sum()
val_dmg_new  = dmg_new_df['지사수수료'].sum()
val_dmg_ret  = dmg_ret_df['지사수수료'].sum()
val_dmg_etc  = dmg_etc_df['지사수수료'].sum()

# 제휴사 랭킹
def comp_rank(src_df, col_rename):
    r = src_df.groupby('제휴사명')['지사수수료'].sum().reset_index().sort_values('지사수수료', ascending=False)
    return r[r['지사수수료'] != 0].rename(columns={'지사수수료': col_rename})

l_new_comp = comp_rank(life_new_df, '익월수수료')
l_ret_comp = comp_rank(life_ret_df, '유지수수료')
d_new_comp = comp_rank(dmg_new_df,  '익월수수료')
d_ret_comp = comp_rank(dmg_ret_df,  '유지수수료')
d_etc_comp = comp_rank(dmg_etc_df,  '기타수수료')


# ==========================================
# 8. 클릭(세션 상태) 감지
# ==========================================
def get_sel(key):
    try:
        rows = st.session_state.get(key, {}).get('selection', {}).get('rows', [])
        return rows[0] if rows else None
    except:
        return None

target_company = None
target_scope   = None

if get_sel('sel_life_new_comp') is not None:
    target_company = l_new_comp.iloc[get_sel('sel_life_new_comp')]['제휴사명']
    target_scope   = '생명 익월'
elif get_sel('sel_life_ret_comp') is not None:
    target_company = l_ret_comp.iloc[get_sel('sel_life_ret_comp')]['제휴사명']
    target_scope   = '생명 유지'
elif get_sel('sel_dmg_new_comp') is not None:
    target_company = d_new_comp.iloc[get_sel('sel_dmg_new_comp')]['제휴사명']
    target_scope   = '손해 익월'
elif get_sel('sel_dmg_ret_comp') is not None:
    target_company = d_ret_comp.iloc[get_sel('sel_dmg_ret_comp')]['제휴사명']
    target_scope   = '손해 유지'
elif get_sel('sel_dmg_etc_comp') is not None:
    target_company = d_etc_comp.iloc[get_sel('sel_dmg_etc_comp')]['제휴사명']
    target_scope   = '손해 기타'

# 상품군 드릴다운
l_new_prod_df = life_new_df[life_new_df['제휴사명'] == target_company] if target_scope == '생명 익월' and target_company else life_new_df
l_new_prod = l_new_prod_df.groupby('상품군')['지사수수료'].sum().reset_index().sort_values('지사수수료', ascending=False)
l_new_prod = l_new_prod[l_new_prod['지사수수료'] > 0]

d_new_prod_df = dmg_new_df[dmg_new_df['제휴사명'] == target_company] if target_scope == '손해 익월' and target_company else dmg_new_df
d_new_prod = d_new_prod_df.groupby('상품군')['지사수수료'].sum().reset_index().sort_values('지사수수료', ascending=False)
d_new_prod = d_new_prod[d_new_prod['지사수수료'] > 0]

d_etc_prod_df = dmg_etc_df[dmg_etc_df['제휴사명'] == target_company] if target_scope == '손해 기타' and target_company else dmg_etc_df
d_etc_prod = d_etc_prod_df.groupby('상품군')['지사수수료'].sum().reset_index().sort_values('지사수수료', ascending=False)
d_etc_prod = d_etc_prod[d_etc_prod['지사수수료'] > 0]


# ==========================================
# 8-1. 업적 대시보드 클릭 감지 추가 (섹션 8보다 먼저 재계산)
# ==========================================
is_jan_new      = life_df['계약일자_정제'].str.startswith('202601', na=False)
is_dmg_date     = damage_df['계약일자_정제'].str.startswith('202601', na=False)
is_jan_new_dmg  = is_dmg_date & is_dmg_new
is_jan_gen_dmg  = is_dmg_date & (damage_df['지급구분'] == '일반')
is_jan_car_dmg  = is_dmg_date & (damage_df['지급구분'] == '자동차')

# 컬럼 없을 때를 대비한 안전한 접근 헬퍼
def safe_sum(df, col):
    return df[col].sum() if col in df.columns else 0

def safe_groupby(df, group_col, val_col):
    if val_col not in df.columns or df.empty:
        return pd.DataFrame(columns=[group_col, val_col])
    r = df.groupby(group_col)[val_col].sum().reset_index().sort_values(val_col, ascending=False)
    return r[r[val_col] > 0]

val_life_hwansan       = safe_sum(life_df[is_jan_new],       '업적지표1')
val_life_premium       = safe_sum(life_df[is_jan_new],       '업적지표2')
val_damage_premium     = safe_sum(damage_df[is_jan_new_dmg], '업적지표3')
val_damage_gen_premium = safe_sum(damage_df[is_jan_gen_dmg], '업적지표3')
val_damage_car_premium = safe_sum(damage_df[is_jan_car_dmg], '업적지표3')

l_h_rank     = safe_groupby(life_df[is_jan_new],       '제휴사명', '업적지표1')
l_p_rank     = safe_groupby(life_df[is_jan_new],       '제휴사명', '업적지표2')
d_p_rank     = safe_groupby(damage_df[is_jan_new_dmg], '제휴사명', '업적지표3')
d_p_gen_rank = safe_groupby(damage_df[is_jan_gen_dmg], '제휴사명', '업적지표3')
d_p_car_rank = safe_groupby(damage_df[is_jan_car_dmg], '제휴사명', '업적지표3')

# 업적 클릭도 target_company/scope에 반영
if get_sel('sel_ach_life1') is not None:
    target_company = l_h_rank.iloc[get_sel('sel_ach_life1')]['제휴사명']
    target_scope   = '생명 업적'
elif get_sel('sel_ach_life2') is not None:
    target_company = l_p_rank.iloc[get_sel('sel_ach_life2')]['제휴사명']
    target_scope   = '생명 업적'
elif get_sel('sel_ach_dmg1') is not None:
    target_company = d_p_rank.iloc[get_sel('sel_ach_dmg1')]['제휴사명']
    target_scope   = '손해 업적'
elif get_sel('sel_ach_dmg2') is not None:
    target_company = d_p_gen_rank.iloc[get_sel('sel_ach_dmg2')]['제휴사명']
    target_scope   = '손해 업적 일반'
elif get_sel('sel_ach_dmg3') is not None:
    target_company = d_p_car_rank.iloc[get_sel('sel_ach_dmg3')]['제휴사명']
    target_scope   = '손해 업적 자동차'

# ==========================================
# 8-2. [상단] 1월 업적 대시보드 (5-Column KPI)
# ==========================================
st.markdown("<h3 style='color:#FFFFFF !important; font-weight: 700; padding-left: 5px; border-left: 5px solid #D4AF37;'>🏆 1월 신계약 및 기타 업적 대시보드 (환산/보험료)</h3>", unsafe_allow_html=True)
col_ach1, col_ach2, col_ach3, col_ach4, col_ach5 = st.columns(5)

with col_ach1:
    st.markdown(f"""
    <div class="kpi-card">
        <p class="kpi-title">생명보험 신계약 환산</p>
        <h2 class="kpi-value">{val_life_hwansan:,.0f} <span class="unit-won">원</span></h2>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<div class='table-toggle-container'><div class='table-toggle-btn'>제휴사별 상세 보기</div></div>", unsafe_allow_html=True)
    disp_dataframe(l_h_rank.rename(columns={'업적지표1': '환산금액'}), hide_index=True, use_container_width=True, height=200, selection_mode="single-row", on_select="rerun", key="sel_ach_life1")

with col_ach2:
    st.markdown(f"""
    <div class="kpi-card">
        <p class="kpi-title">생명보험 신계약 보험료</p>
        <h2 class="kpi-value">{val_life_premium:,.0f} <span class="unit-won">원</span></h2>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<div class='table-toggle-container'><div class='table-toggle-btn'>제휴사별 상세 보기</div></div>", unsafe_allow_html=True)
    disp_dataframe(l_p_rank.rename(columns={'업적지표2': '보험료'}), hide_index=True, use_container_width=True, height=200, selection_mode="single-row", on_select="rerun", key="sel_ach_life2")

with col_ach3:
    st.markdown(f"""
    <div class="kpi-card">
        <p class="kpi-title">손해보험 신계약 보험료</p>
        <h2 class="kpi-value">{val_damage_premium:,.0f} <span class="unit-won">원</span></h2>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<div class='table-toggle-container'><div class='table-toggle-btn'>제휴사별 상세 보기</div></div>", unsafe_allow_html=True)
    disp_dataframe(d_p_rank.rename(columns={'업적지표3': '총보험료'}), hide_index=True, use_container_width=True, height=200, selection_mode="single-row", on_select="rerun", key="sel_ach_dmg1")

with col_ach4:
    st.markdown(f"""
    <div class="kpi-card">
        <p class="kpi-title">손해보험 일반 보험료</p>
        <h2 class="kpi-value">{val_damage_gen_premium:,.0f} <span class="unit-won">원</span></h2>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<div class='table-toggle-container'><div class='table-toggle-btn'>제휴사별 상세 보기</div></div>", unsafe_allow_html=True)
    disp_dataframe(d_p_gen_rank.rename(columns={'업적지표3': '총보험료'}), hide_index=True, use_container_width=True, height=200, selection_mode="single-row", on_select="rerun", key="sel_ach_dmg2")

with col_ach5:
    st.markdown(f"""
    <div class="kpi-card">
        <p class="kpi-title">손해보험 자동차 보험료</p>
        <h2 class="kpi-value">{val_damage_car_premium:,.0f} <span class="unit-won">원</span></h2>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<div class='table-toggle-container'><div class='table-toggle-btn'>제휴사별 상세 보기</div></div>", unsafe_allow_html=True)
    disp_dataframe(d_p_car_rank.rename(columns={'업적지표3': '총보험료'}), hide_index=True, use_container_width=True, height=200, selection_mode="single-row", on_select="rerun", key="sel_ach_dmg3")

st.markdown("---")

# ==========================================
# 8-3. [중간] 업적 전용 상세데이터
# ==========================================
st.markdown("<h3 style='color:#FFFFFF !important; font-weight: 700; padding-left: 5px; border-left: 5px solid #D4AF37;'>📊 1월 전체 업적 전용 상세데이터</h3>", unsafe_allow_html=True)

if target_scope == '생명 업적':
    ach_detail_df = life_df[is_jan_new].copy(); ach_scope = "생명 업적"
elif target_scope == '손해 업적':
    ach_detail_df = damage_df[is_jan_new_dmg].copy(); ach_scope = "손해 업적"
elif target_scope == '손해 업적 일반':
    ach_detail_df = damage_df[is_jan_gen_dmg].copy(); ach_scope = "손해 업적 일반"
elif target_scope == '손해 업적 자동차':
    ach_detail_df = damage_df[is_jan_car_dmg].copy(); ach_scope = "손해 업적 자동차"
else:
    ach_detail_df = pd.concat([life_df[is_jan_new], damage_df[is_jan_new_dmg], damage_df[is_jan_gen_dmg], damage_df[is_jan_car_dmg]])
    ach_scope = "1월 업적 통합"

if target_scope in ['생명 업적', '손해 업적', '손해 업적 일반', '손해 업적 자동차'] and target_company:
    st.markdown(f"<h4 style='color:#D4AF37 !important; font-weight: 700;'>🔎 조회중: [{target_company}] ({ach_scope}) 상세 데이터</h4>", unsafe_allow_html=True)
    ach_detail_df = ach_detail_df[ach_detail_df['제휴사명'] == target_company]
else:
    st.markdown("<h4 style='color:#FFFFFF !important; font-weight: 700;'>📄 1월 업적 데이터(신계약/기타) 전체 상세</h4>", unsafe_allow_html=True)
    st.caption("💡 위 업적 표에서 회사를 클릭하면 이 자리에 해당 회사의 상세 내역이 연동됩니다.")

display_cols_ach = ['계약일자_정제', '제휴사명', 'FC명', '지급구분', '상품군', '상품명', '계약자', '업적지표2', '업적지표3', '업적지표1']
ach_display_df = ach_detail_df[[c for c in display_cols_ach if c in ach_detail_df.columns]].copy()
ach_display_df.rename(columns={'계약일자_정제': '계약일자', '업적지표1': '환산', '업적지표2': '생명_보험료', '업적지표3': '손해_보험료'}, inplace=True)
if ach_scope == '생명 업적' and '손해_보험료' in ach_display_df.columns:
    ach_display_df = ach_display_df.drop(columns=['손해_보험료'])
elif '손해' in ach_scope:
    if '생명_보험료' in ach_display_df.columns: ach_display_df = ach_display_df.drop(columns=['생명_보험료'])
    if '환산' in ach_display_df.columns: ach_display_df = ach_display_df.drop(columns=['환산'])
disp_dataframe(ach_display_df, use_container_width=True, height=300)
st.markdown("---")

# ==========================================
# 9. 전체 수수료 대시보드 (더블 컨테이너)
# ==========================================
st.markdown("<h3 style='color:#FFFFFF !important; font-weight: 700; padding-left: 5px; border-left: 5px solid #D4AF37;'>💰 전체 수수료 대시보드 (지사수수료 기준)</h3>", unsafe_allow_html=True)

zone_left, _sp, zone_right = st.columns([1, 0.03, 1])

with zone_left:
    st.markdown(f'''
    <div class="kpi-card-zone" style="margin-bottom: 24px;">
        <p class="zone-title">🔵 생명보험 총 수수료</p>
        <h2 class="zone-value">{val_life_new + val_life_ret:,.0f} <span class="unit-won">원</span></h2>
        <p style="color:#A0AEC0; margin-top:5px; font-size:1.1rem;">익월(신계약) + 유지(환수 포함) 수수료 합산</p>
    </div>
    ''', unsafe_allow_html=True)

    sub_l1, sub_l2 = st.columns(2)
    with sub_l1:
        st.markdown(f"<div class='kpi-card-zone'><p class='zone-title'>① 생명 익월</p><h2 class='zone-value-sub'>{val_life_new:,.0f} <span class='unit-won'>원</span></h2></div>", unsafe_allow_html=True)
        st.markdown("<div class='table-toggle-container'><div class='table-toggle-btn'>제휴사별 상세 보기</div></div>", unsafe_allow_html=True)
        disp_dataframe(l_new_comp, hide_index=True, use_container_width=True, height=200,
                       selection_mode="single-row", on_select="rerun", key="sel_life_new_comp")
        if target_scope == '생명 익월' and target_company and len(l_new_prod) > 0:
            st.markdown(f"<div class='table-toggle-container'><div class='table-toggle-btn'>[{target_company}] 상품군별 상세 보기</div></div>", unsafe_allow_html=True)
            disp_dataframe(l_new_prod, hide_index=True, use_container_width=True, height=150)

    with sub_l2:
        st.markdown(f"<div class='kpi-card-zone'><p class='zone-title'>② 생명 유지(환수포함)</p><h2 class='zone-value-sub'>{val_life_ret:,.0f} <span class='unit-won'>원</span></h2></div>", unsafe_allow_html=True)
        st.markdown("<div class='table-toggle-container'><div class='table-toggle-btn'>제휴사별 상세 보기</div></div>", unsafe_allow_html=True)
        disp_dataframe(l_ret_comp, hide_index=True, use_container_width=True, height=200,
                       selection_mode="single-row", on_select="rerun", key="sel_life_ret_comp")

with zone_right:
    st.markdown(f'''
    <div class="kpi-card-zone" style="margin-bottom: 24px;">
        <p class="zone-title">🟠 손해보험 총 수수료</p>
        <h2 class="zone-value">{val_dmg_new + val_dmg_ret + val_dmg_etc:,.0f} <span class="unit-won">원</span></h2>
        <p style="color:#A0AEC0; margin-top:5px; font-size:1.1rem;">익월(신계약) + 유지 + 기타(일반/차/환수) 합산</p>
    </div>
    ''', unsafe_allow_html=True)

    sub_r1, sub_r2, sub_r3 = st.columns(3)
    with sub_r1:
        st.markdown(f"<div class='kpi-card-zone' style='padding: 24px 10px 10px 10px !important;'><p class='zone-title' style='font-size:1.2rem;'>③ 손해 익월</p><h2 class='zone-value-sub' style='font-size:2.4rem;'>{val_dmg_new:,.0f} <span class='unit-won'>원</span></h2></div>", unsafe_allow_html=True)
        st.markdown("<div class='table-toggle-container'><div class='table-toggle-btn'>상세 보기</div></div>", unsafe_allow_html=True)
        disp_dataframe(d_new_comp, hide_index=True, use_container_width=True, height=200,
                       selection_mode="single-row", on_select="rerun", key="sel_dmg_new_comp")
        if target_scope == '손해 익월' and target_company and len(d_new_prod) > 0:
            st.markdown(f"<div class='table-toggle-container'><div class='table-toggle-btn'>[{target_company[:4]}] 상품군</div></div>", unsafe_allow_html=True)
            disp_dataframe(d_new_prod, hide_index=True, use_container_width=True, height=150)

    with sub_r2:
        st.markdown(f"<div class='kpi-card-zone' style='padding: 24px 10px 10px 10px !important;'><p class='zone-title' style='font-size:1.2rem;'>④ 손해 유지</p><h2 class='zone-value-sub' style='font-size:2.4rem;'>{val_dmg_ret:,.0f} <span class='unit-won'>원</span></h2></div>", unsafe_allow_html=True)
        st.markdown("<div class='table-toggle-container'><div class='table-toggle-btn'>상세 보기</div></div>", unsafe_allow_html=True)
        disp_dataframe(d_ret_comp, hide_index=True, use_container_width=True, height=200,
                       selection_mode="single-row", on_select="rerun", key="sel_dmg_ret_comp")

    with sub_r3:
        st.markdown(f"<div class='kpi-card-zone' style='padding: 24px 10px 10px 10px !important;'><p class='zone-title' style='font-size:1.2rem;'>⑤ 손해 기타</p><h2 class='zone-value-sub' style='font-size:2.4rem;'>{val_dmg_etc:,.0f} <span class='unit-won'>원</span></h2></div>", unsafe_allow_html=True)
        st.markdown("<div class='table-toggle-container'><div class='table-toggle-btn'>상세 보기</div></div>", unsafe_allow_html=True)
        disp_dataframe(d_etc_comp, hide_index=True, use_container_width=True, height=200,
                       selection_mode="single-row", on_select="rerun", key="sel_dmg_etc_comp")
        if target_scope == '손해 기타' and target_company and len(d_etc_prod) > 0:
            st.markdown(f"<div class='table-toggle-container'><div class='table-toggle-btn'>[{target_company[:4]}] 상품군</div></div>", unsafe_allow_html=True)
            disp_dataframe(d_etc_prod, hide_index=True, use_container_width=True, height=150)


# ==========================================
# 10. 수수료 전용 상세 데이터
# ==========================================
st.markdown("---")
st.markdown("<h3 style='color:#FFFFFF !important; font-weight: 700; padding-left: 5px; border-left: 5px solid #D4AF37;'>📊 수수료 전용 상세데이터</h3>", unsafe_allow_html=True)

if target_scope == '생명 익월':    comm_detail_df = life_new_df.copy()
elif target_scope == '생명 유지':  comm_detail_df = life_ret_df.copy()
elif target_scope == '손해 익월':  comm_detail_df = dmg_new_df.copy()
elif target_scope == '손해 유지':  comm_detail_df = dmg_ret_df.copy()
elif target_scope == '손해 기타':  comm_detail_df = dmg_etc_df.copy()
else:                              comm_detail_df = filtered_df.copy()

if target_company:
    comm_detail_df = comm_detail_df[comm_detail_df['제휴사명'] == target_company]
    st.markdown(f"<h4 style='color:#D4AF37 !important; font-weight: 700;'>🔎 조회중: [{target_company}] ({target_scope or '전체'}) 상세 데이터</h4>", unsafe_allow_html=True)
else:
    st.caption("💡 위 수수료 표에서 제휴사를 클릭하면 이 자리에 상세 내역이 연동됩니다.")

display_cols = ['계약일자_정제', '제휴사명', 'FC명', '지급구분', '상품군', '상품명', '계약자', '지사수수료']
comm_display_df = comm_detail_df[[c for c in display_cols if c in comm_detail_df.columns]].copy()
comm_display_df.rename(columns={'계약일자_정제': '계약일자', '지사수수료': '지사수수료(원)'}, inplace=True)

disp_dataframe(comm_display_df, use_container_width=True, height=400)
