import pandas as pd
import json

def load_data():
    file_path = r"G:\내 드라이브\안티그래비티\TEST\수수료분석\최종병합_수수료명세서_압축판.xlsx"
    df_life = pd.read_excel(file_path, sheet_name='생명보험사_202602')
    df_damage = pd.read_excel(file_path, sheet_name='손해보험사_202602')
    
    df_life['보험군'] = '생명보험'
    df_damage['보험군'] = '손해보험'
    
    # 원본 컬럼 유지 (강제 업적지표 변환 제거)
    # 손해보험의 '수정보험료'가 있으면 '보험료'로 통일하여 다루기 쉽게 함
    if '수정보험료' in df_damage.columns and '보험료' not in df_damage.columns:
        df_damage['보험료'] = df_damage['수정보험료']
        
    common_cols = ['보험군', '제휴사명', 'FC명', '계약자', '상품군', '상품명', '증권번호', '계약일자', '지급구분',
                   '환산성적', '보험료', 'FC수수료', '지사수수료']
    val_cols = [c for c in common_cols if c in df_life.columns or c in df_damage.columns]
    
    df_all = pd.concat([df_life[df_life.columns.intersection(val_cols)], 
                        df_damage[df_damage.columns.intersection(val_cols)]], ignore_index=True)
    
    df_all['계약일자_정제'] = df_all['계약일자'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip() if '계약일자' in df_all.columns else '00000000'
    
    for c in ['환산성적', '보험료', 'FC수수료', '지사수수료']:
        if c in df_all.columns: df_all[c] = pd.to_numeric(df_all[c], errors='coerce').fillna(0)
            
    for c in ['증권번호', '계약일자_정제', '계약자', 'FC명', '제휴사명', '상품명', '상품군', '지급구분']:
        if c in df_all.columns: df_all[c] = df_all[c].astype(str).fillna('-')
            
    df_all['상품군'] = df_all['상품군'].replace('nan', '분류없음').str.strip()
    df_all['FC명'] = df_all['FC명'].replace('nan', '알수없음')
    df_all['제휴사명'] = df_all['제휴사명'].replace('nan', '알수없음')
    return df_all

df = load_data()

life_df = df[df['보험군'] == '생명보험']
damage_df = df[df['보험군'] == '손해보험']

is_life_new = life_df['지급구분'].str.contains('신계약', na=False)
is_dmg_new = damage_df['지급구분'].str.contains('신계약', na=False)

is_jan_new = life_df['계약일자_정제'].str.startswith('202601', na=False)
is_dmg_date_match = damage_df['계약일자_정제'].str.startswith('202601', na=False)

life_new_df = life_df[is_life_new]
life_ret_df = life_df[~is_life_new]
dmg_new_df = damage_df[is_dmg_new]
dmg_ret_df = damage_df[damage_df['지급구분'].str.contains('유지', na=False)]
dmg_etc_df = damage_df[damage_df['지급구분'].str.contains('일반|자동차|환수', na=False, regex=True)]
dmg_gen_only_df = damage_df[damage_df['지급구분'].str.contains('일반', na=False)]

def get_kpi(df_src, cond, col):
    return df_src[cond][col].sum() if cond is not None else df_src[col].sum()

val_life_hwansan = get_kpi(life_df, is_jan_new, '환산성적') if '환산성적' in life_df.columns else 0
val_life_premium = get_kpi(life_df, is_jan_new, '보험료') if '보험료' in life_df.columns else 0
val_dmg_premium = get_kpi(damage_df, is_dmg_date_match & is_dmg_new, '보험료') if '보험료' in damage_df.columns else 0
val_dmg_gen_premium = get_kpi(damage_df, is_dmg_date_match & (damage_df['지급구분']=='일반'), '보험료') if '보험료' in damage_df.columns else 0
val_dmg_car_premium = get_kpi(damage_df, is_dmg_date_match & (damage_df['지급구분']=='자동차'), '보험료') if '보험료' in damage_df.columns else 0

val_life_new = life_new_df['지사수수료'].sum()
val_life_ret = life_ret_df['지사수수료'].sum()
val_dmg_new = dmg_new_df['지사수수료'].sum()
val_dmg_ret = dmg_ret_df['지사수수료'].sum()
val_dmg_etc = dmg_etc_df['지사수수료'].sum()

def get_rank(df_src, val_col):
    df_grp = df_src.groupby('제휴사명')[val_col].sum().reset_index()
    return df_grp[df_grp[val_col] > 0].sort_values(val_col, ascending=False)

def to_html_table(df_grp, val_col, target_scope):
    if df_grp.empty: return "<p style='color:#64748B;'>데이터 없음</p>"
    html = f'''<div class="table-toggle-btn" onclick="openCompModal('{target_scope}')">제휴사별 상세 보기 ▼</div>'''
    html += '<div class="table-container"><table>'
    for _, row in df_grp.iterrows():
        comp, val = row['제휴사명'], row[val_col]
        color_st = 'color:#E63946;' if val < 0 else ''
        html += f'<tr onclick="openDetailModal(&quot;{target_scope}&quot;, &quot;{comp}&quot;, null)" style="cursor:pointer;"><td style="text-align:left;">{comp}</td><td style="text-align:right; {color_st}">{val:,.0f}</td></tr>'
    html += '</table></div>'
    return html

def to_top10_html(title, color, scope_df, target_scope):
    html = f'<div class="top5-card" style="border-top-color:{color};"><div class="top5-title">{title}</div>'
    if scope_df.empty: return html + "<p style='color:#64748B;'>데이터 없음</p></div>"
    
    top10 = scope_df.groupby('FC명')['지사수수료'].sum().reset_index().sort_values('지사수수료', ascending=False).head(10)
    medals = ["🥇", "🥈", "🥉", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣", "9️⃣", "🔟"]
    colors = ["#D4AF37", "#C0C0C0", "#CD7F32", "#94A3B8", "#94A3B8", "#94A3B8", "#94A3B8", "#94A3B8", "#94A3B8", "#94A3B8"]
    
    for i, row in top10.reset_index(drop=True).iterrows():
        man_val, fc = int(row['지사수수료'] / 10000), row['FC명']
        bg = "rgba(212,175,55,0.08)" if i == 0 else "rgba(255,255,255,0.02)"
        border = "1px solid rgba(212,175,55,0.25)" if i == 0 else "1px solid rgba(255,255,255,0.04)"
        icon = medals[i] if i < len(medals) else f"{i+1}"
        html += f'''
        <button class="top10-row" style="background:{bg}; border:{border};" onclick="openDetailModal(&quot;{target_scope}&quot;, null, &quot;{fc}&quot;)">
            <span style="font-size:1.1rem; width:25px; text-align:center;">{icon}</span>
            <span class="fc-name" style="color:{colors[i] if i < 10 else '#94A3B8'};">{fc}</span>
            <span class="fc-amt">{man_val:,}<span class="unit-won-sm">만원</span></span>
        </button>'''
    return html + '</div>'

export_cols = ['보험군', '제휴사명', 'FC명', '계약자', '상품군', '상품명', '계약일자_정제', '지급구분', '환산성적', '보험료', '지사수수료']
# 컬럼 없으면 0으로 채우기
for c in ['환산성적', '보험료']:
    if c not in df.columns: df[c] = 0
export_cols = [c for c in export_cols if c in df.columns]
df_json = df[export_cols].rename(columns={'계약일자_정제': '계약일자'})
json_data = df_json.to_json(orient='records', force_ascii=False)

html_content = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>웰스FA VVIP 수수료 정밀분석</title>
    <style>
        :root {{ 
            --bg-main: #0A1128; 
            --bg-card: #16203B; 
            --color-gold: #D4AF37; 
            --color-text: #E2E8F0; 
        }}
        
        body {{ 
            background: var(--bg-main); 
            color: var(--color-text); 
            font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif; 
            margin: 0; 
            padding: 20px; 
            line-height: 1.6;
        }}

        #passwordGate {{ 
            position: fixed; top:0; left:0; width:100%; height:100%; 
            background: var(--bg-main); z-index: 9999; 
            display: flex; flex-direction: column; align-items: center; justify-content: center; 
        }}
        .pass-input {{ 
            background: var(--bg-card); color: var(--color-gold); 
            border: 1px solid var(--color-gold); padding: 15px; 
            font-size: 1.5rem; text-align: center; border-radius: 8px; 
            margin-bottom: 20px; outline: none; width: 300px;
        }}
        .pass-btn {{ 
            background: var(--color-gold); color: #000; 
            font-weight: 800; font-size: 1.2rem; padding: 12px 40px; 
            border: none; border-radius: 8px; cursor: pointer; transition: all 0.2s;
        }}
        .pass-btn:hover {{ filter: brightness(1.2); transform: scale(1.05); }}

        #mainDashboard {{ display: none; opacity: 0; transition: opacity 0.5s ease-in; }}
        
        .container {{ max-width: 1600px; margin: 0 auto; }}
        
        h1 {{ 
            font-size: 4.0rem; font-weight: 900; 
            background: linear-gradient(90deg, #BF953F, #FCF6BA, #D4AF37, #FBF5B7, #AA771C); 
            -webkit-background-clip: text; -webkit-text-fill-color: transparent; 
            filter: drop-shadow(0px 4px 10px rgba(212,175,55,0.4)); 
            margin: 0 0 10px 0; letter-spacing: -1px;
        }}
        
        hr {{ border: 0; border-top: 1px solid rgba(255,255,255,0.05); margin: 40px 0; }}
        
        .section-title {{ 
            color:#FFFFFF !important; font-weight: 900; font-size: 1.8rem;
            padding-left: 15px; border-left: 6px solid var(--color-gold); 
            margin-bottom: 25px; letter-spacing: -0.5px;
        }}

        .grid-5 {{ display: grid; grid-template-columns: repeat(5, 1fr); gap: 20px; }}
        
        /* KPI Cards */
        .kpi-card {{
            background: linear-gradient(180deg, var(--bg-card) 0%, #0A1128 100%);
            border-radius: 16px; border: 1px solid rgba(255,255,255,0.05);
            border-top: 5px solid var(--color-gold);
            padding: 30px 24px; text-align: center;
            box-shadow: 0 10px 30px rgba(0,0,0,0.4);
        }}
        .kpi-title {{ color: #A0AEC0; font-weight:700; font-size:1.3rem; margin:0 0 15px 0; }}
        .kpi-value-container {{ 
            display: flex; align-items: baseline; justify-content: center; 
            gap: 4px; overflow: hidden; white-space: nowrap; 
        }}
        .kpi-value-main {{ 
            color: var(--color-gold); font-weight:900; font-size: clamp(1.8rem, 5vw, 3.2rem); 
            margin:0; line-height:1.1; letter-spacing: -1px;
        }}
        .unit-won {{ font-size: 1.2rem; color: #94A3B8; font-weight: 700; }}

        /* Total Section Styles */
        .total-banner {{
            background: linear-gradient(135deg, #1A264D 0%, #0A1128 100%);
            border-radius: 20px; padding: 25px 40px; margin-bottom: 30px;
            border: 2px solid var(--color-gold);
            box-shadow: 0 15px 40px rgba(0,0,0,0.5);
            display: flex; align-items: center; justify-content: space-between;
        }}
        .total-label {{ font-size: 1.5rem; color: #FFF; font-weight: 800; }}
        .total-value-box {{ text-align: right; }}
        .total-value {{ font-size: 4.5rem; font-weight: 950; color: var(--color-gold); line-height: 1; }}

        /* Double Zone Layout */
        .double-zone-wrapper {{ display: flex; gap: 30px; margin-top:30px; }}
        .zone-box {{ 
            flex: 1; background: rgba(255, 255, 255, 0.02); 
            border-radius: 20px; padding: 30px; 
            border: 1px solid rgba(255, 255, 255, 0.05);
            box-shadow: 0 15px 45px rgba(0,0,0,0.5);
        }}
        .zone-header-card {{
            background: linear-gradient(180deg, var(--bg-card) 0%, #0A1128 100%);
            border-radius: 16px; padding: 40px 30px; 
            border-top: 5px solid var(--color-gold); 
            text-align: center; margin-bottom: 25px;
            border-left: 1px solid rgba(255, 255, 255, 0.05);
            border-right: 1px solid rgba(255, 255, 255, 0.05);
            border-bottom: 1px solid rgba(255, 255, 255, 0.05);
        }}
        .zone-title {{ color: #A0AEC0 !important; font-weight:800; font-size:1.6rem; margin:0 0 10px 0; }}
        .zone-value-container {{ display: flex; align-items: baseline; justify-content: center; gap: 8px; }}
        .zone-value {{ color: var(--color-gold) !important; font-weight:900; font-size: 4.2rem; margin:0; line-height: 1; }}
        
        .zone-inner-grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
        .zone-inner-grid-3 {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 15px; }}
        
        .sub-kpi-card {{
            background: var(--bg-card); border-radius: 14px;
            padding: 25px 20px; text-align: center;
            border: 1px solid rgba(255, 255, 255, 0.05);
            border-top: 3px solid var(--color-gold);
        }}
        .sub-kpi-title {{ color: #A0AEC0; font-weight:700; font-size:1.2rem; margin-bottom: 8px; }}
        .sub-kpi-value-container {{ display: flex; align-items: baseline; justify-content: center; gap: 4px; }}
        .sub-kpi-value {{ color: var(--color-gold); font-weight:900; font-size:2.8rem; margin:0; }}

        /* VVIP Table Styles */
        .table-toggle-btn {{ 
            background: rgba(212,175,55,0.1); border: 1px solid rgba(212,175,55,0.2); 
            border-radius: 8px; padding: 10px; color: var(--color-gold); 
            cursor: pointer; text-align: center; font-weight: 700; 
            margin: 15px 0 10px 0; font-size:1rem; transition: all 0.2s;
        }}
        .table-toggle-btn:hover {{ background: var(--color-gold); color: #000; }}
        
        .table-container {{ max-height: 250px; overflow-y: auto; margin-top: 10px; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th, td {{ 
            padding: 12px 10px; 
            border-bottom: 1px solid rgba(255, 255, 255, 0.05); 
            font-size: 1.05rem;
        }}
        th {{ color: #A0AEC0; font-weight: 800; border-bottom: 1px solid rgba(255, 255, 255, 0.1); }}
        tr:hover td {{ background: rgba(255, 255, 255, 0.03); }}
        
        /* FC Cards (Top 10) */
        .top5-card {{ 
            background: linear-gradient(180deg, #16203B, #0A1128); 
            border-radius: 16px; border: 1px solid rgba(255,255,255,0.05); 
            border-top: 5px solid; padding: 20px;
        }}
        .top5-title {{ color: #FFFFFF; font-weight:900; font-size:1.3rem; margin-bottom:15px; }}
        .top10-row {{ 
            display:flex; align-items:center; justify-content:space-between; 
            border-radius:10px; padding: 10px 12px; margin-bottom:6px; 
            cursor: pointer; transition: all 0.2s; border: 1px solid rgba(255,255,255,0.03);
            width: 100%; text-align: left; outline: none;
        }}
        .top10-row:hover {{ transform: translateX(5px); background: rgba(255,255,255,0.05); }}
        .fc-name {{ flex:1; margin-left:10px; font-weight:800; font-size:1.1rem; color: #FFF; }}
        .fc-amt {{ color:var(--color-gold); font-weight:900; font-size:1.05rem; }}

        /* Modal Styles (Clean Fintech) */
        #modalOverlay {{ 
            display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; 
            background: rgba(0,0,0,0.85); backdrop-filter: blur(5px);
            z-index: 10000; justify-content: center; align-items: center; 
        }}
        #modalContent {{ 
            background: #0A1128; border: 1px solid var(--color-gold); 
            border-radius: 20px; width: 95%; max-width: 1300px; max-height: 90vh; 
            display: flex; flex-direction: column; overflow: hidden;
            box-shadow: 0 25px 60px rgba(0,0,0,0.8);
        }}
        .modal-header {{ 
            padding: 30px; border-bottom: 1px solid rgba(255,255,255,0.1); 
            display: flex; justify-content: space-between; align-items: center; 
            background: #16203B;
        }}
        .modal-body {{ padding: 30px; overflow-y: auto; flex: 1; }}
        .modal-close {{ background: none; border: none; color: #64748B; font-size: 2.5rem; cursor: pointer; line-height: 1; }}
        .modal-close:hover {{ color: #FFF; }}
        
        #modalTable th {{ 
            position: sticky; top: -1px; background: #16203B; color: var(--color-gold); 
            padding: 15px; border-bottom: 2px solid var(--color-gold); z-index: 10;
        }}
        #modalTable td {{ padding: 12px; border-bottom: 1px solid rgba(255,255,255,0.05); text-align: center; }}
        
        .modal-summary-grid {{ display:flex; gap:25px; margin-bottom:30px; }}
        .metric-box {{ 
            background: #16203B; border-radius: 15px; padding: 25px; 
            text-align: center; flex: 1; border: 1px solid rgba(255,255,255,0.05); 
            box-shadow: 0 5px 15px rgba(0,0,0,0.2);
        }}
        
        ::-webkit-scrollbar {{ width: 8px; }}
        ::-webkit-scrollbar-track {{ background: transparent; }}
        ::-webkit-scrollbar-thumb {{ background: rgba(255,255,255,0.1); border-radius: 10px; }}
        ::-webkit-scrollbar-thumb:hover {{ background: var(--color-gold); }}
    </style>
</head>
<body>

<div id="passwordGate">
    <h1 style="font-size:3.5rem; margin-bottom: 15px;">웰스FA SECURITY</h1>
    <p style="color:#A0AEC0; margin-bottom:40px; font-size:1.2rem;">마스터피스 대시보드 접근 권한 확인이 필요합니다.</p>
    <input type="password" id="passInput" class="pass-input" placeholder="ACCESS CODE">
    <button class="pass-btn" onclick="checkPass()">인증 및 입장</button>
</div>

<div id="mainDashboard" class="container">
    <h1>웰스FA 26년 2월 수수료 정밀분석</h1>
    <p style="color:#94A3B8; font-size:1.1rem; margin-bottom: 25px;">※ 1월 영업 업적 기준 정교화 데이터</p>
    <hr>
    
    <div class="section-title">🏆 1월 신계약 및 기타 업적 대시보드 (환산/보험료)</div>
    <div class="grid-5">
        <div class="kpi-card">
            <p class="kpi-title">생명보험 신계약 환산</p>
            <div class="kpi-value-container">
                <p class="kpi-value-main">{val_life_hwansan:,.0f}</p><span class="unit-won">원</span>
            </div>
            {to_html_table(get_rank(life_df[is_jan_new], '환산성적') if '환산성적' in life_df else pd.DataFrame(), '환산성적', 'ach_life_h')}
        </div>
        <div class="kpi-card">
            <p class="kpi-title">생명보험 신계약 보험료</p>
            <div class="kpi-value-container">
                <p class="kpi-value-main">{val_life_premium:,.0f}</p><span class="unit-won">원</span>
            </div>
            {to_html_table(get_rank(life_df[is_jan_new], '보험료') if '보험료' in life_df else pd.DataFrame(), '보험료', 'ach_life_p')}
        </div>
        <div class="kpi-card">
            <p class="kpi-title">손해보험 신계약 보험료</p>
            <div class="kpi-value-container">
                <p class="kpi-value-main">{val_dmg_premium:,.0f}</p><span class="unit-won">원</span>
            </div>
            {to_html_table(get_rank(damage_df[is_dmg_date_match & is_dmg_new], '보험료') if '보험료' in damage_df else pd.DataFrame(), '보험료', 'ach_dmg_new')}
        </div>
        <div class="kpi-card">
            <p class="kpi-title">손해보험 일반 보험료</p>
            <div class="kpi-value-container">
                <p class="kpi-value-main">{val_dmg_gen_premium:,.0f}</p><span class="unit-won">원</span>
            </div>
            {to_html_table(get_rank(damage_df[is_dmg_date_match & (damage_df['지급구분']=='일반')], '보험료') if '보험료' in damage_df else pd.DataFrame(), '보험료', 'ach_dmg_gen')}
        </div>
        <div class="kpi-card">
            <p class="kpi-title">손해보험 자동차 보험료</p>
            <div class="kpi-value-container">
                <p class="kpi-value-main">{val_dmg_car_premium:,.0f}</p><span class="unit-won">원</span>
            </div>
            {to_html_table(get_rank(damage_df[is_dmg_date_match & (damage_df['지급구분']=='자동차')], '보험료') if '보험료' in damage_df else pd.DataFrame(), '보험료', 'ach_dmg_car')}
        </div>
    </div>
    
    <hr>
    <div class="section-title">💰 전체 수수료 대시보드 (지사수수료 기준)</div>
    
    <div class="total-banner">
        <div class="total-label">🏢 웰스FA 지사 총 수수료 합계 <span style="font-size:0.9rem; font-weight:400; opacity:0.7; display:block;">(생명 + 손해 전체 통합)</span></div>
        <div class="total-value-box">
            <span class="total-value">{val_life_new + val_life_ret + val_dmg_new + val_dmg_ret + val_dmg_etc:,.0f}</span>
            <span class="unit-won" style="font-size:2.0rem; margin-left:10px;">원</span>
        </div>
    </div>

    <div class="double-zone-wrapper">
        <div class="zone-box">
            <div class="zone-header-card">
                <p class="zone-title">🔵 생명보험 총 수수료</p>
                <div class="zone-value-container">
                    <p class="zone-value">{val_life_new + val_life_ret:,.0f}</p>
                    <span class="unit-won" style="font-size:1.5rem;">원</span>
                </div>
                <p style="color:#A0AEC0; margin-top:10px; font-size:1.1rem;">익월(신계약) + 유지(환수 포함) 수수료 합산</p>
            </div>
            <div class="zone-inner-grid-2">
                <div class="sub-kpi-card">
                    <p class="sub-kpi-title">① 생명 익월 수수료</p>
                    <div class="sub-kpi-value-container">
                        <p class="sub-kpi-value">{val_life_new:,.0f}</p><span class="unit-won" style="font-size:1.2rem; margin-left:4px;">원</span>
                    </div>
                    {to_html_table(get_rank(life_new_df, '지사수수료'), '지사수수료', 'fee_life_new')}
                </div>
                <div class="sub-kpi-card">
                    <p class="sub-kpi-title">② 생명 유지</p>
                    <div class="sub-kpi-value-container">
                        <p class="sub-kpi-value">{val_life_ret:,.0f}</p><span class="unit-won" style="font-size:1.2rem; margin-left:4px;">원</span>
                    </div>
                    {to_html_table(get_rank(life_ret_df, '지사수수료'), '지사수수료', 'fee_life_ret')}
                </div>
            </div>
        </div>
        
        <div class="zone-box">
            <div class="zone-header-card">
                <p class="zone-title">🟠 손해보험 총 수수료</p>
                <div class="zone-value-container">
                    <p class="zone-value">{val_dmg_new + val_dmg_ret + val_dmg_etc:,.0f}</p>
                    <span class="unit-won" style="font-size:1.5rem;">원</span>
                </div>
                <p style="color:#A0AEC0; margin-top:10px; font-size:1.1rem;">익월(신계약) + 유지 + 기타(일반/차/환수) 합산</p>
            </div>
            <div class="zone-inner-grid-3">
                <div class="sub-kpi-card">
                    <p class="sub-kpi-title">③ 손해 익월</p>
                    <div class="sub-kpi-value-container">
                        <p class="sub-kpi-value" style="font-size:2.2rem;">{val_dmg_new:,.0f}</p><span class="unit-won" style="font-size:1.0rem; margin-left:2px;">원</span>
                    </div>
                    {to_html_table(get_rank(dmg_new_df, '지사수수료'), '지사수수료', 'fee_dmg_new')}
                </div>
                <div class="sub-kpi-card">
                    <p class="sub-kpi-title">④ 손해 유지</p>
                    <div class="sub-kpi-value-container">
                        <p class="sub-kpi-value" style="font-size:2.2rem;">{val_dmg_ret:,.0f}</p><span class="unit-won" style="font-size:1.0rem; margin-left:2px;">원</span>
                    </div>
                    {to_html_table(get_rank(dmg_ret_df, '지사수수료'), '지사수수료', 'fee_dmg_ret')}
                </div>
                <div class="sub-kpi-card">
                    <p class="sub-kpi-title">⑤ 손해 기타</p>
                    <div class="sub-kpi-value-container">
                        <p class="sub-kpi-value" style="font-size:2.2rem;">{val_dmg_etc:,.0f}</p><span class="unit-won" style="font-size:1.0rem; margin-left:2px;">원</span>
                    </div>
                    {to_html_table(get_rank(dmg_etc_df, '지사수수료'), '지사수수료', 'fee_dmg_etc')}
                </div>
            </div>
        </div>
    </div>
    
    <hr>
    <div class="section-title">👑 부문별 지사수수료 기여 Top 10 FC</div>
    <div class="grid-5">
        {to_top10_html('🔵 생명 익월', '#3B82F6', life_new_df, 'fee_life_new')}
        {to_top10_html('🟢 생명 유지', '#10B981', life_ret_df, 'fee_life_ret')}
        {to_top10_html('🟠 손해 익월', '#F59E0B', dmg_new_df, 'fee_dmg_new')}
        {to_top10_html('🟣 손해 유지', '#8B5CF6', dmg_ret_df, 'fee_dmg_ret')}
        {to_top10_html('⚪ 손해 기타', '#94A3B8', dmg_gen_only_df, 'fee_dmg_gen')}
    </div>
    
    <div style="text-align:center; color:#334155; font-size:0.9rem; margin:100px 0 50px 0; padding:40px; border-top:1px solid rgba(255,255,255,0.03);">
        웰스FA 마스터피스 대시보드 · <span style="color:var(--color-gold);">Supreme Edition</span> · Powered by Antigravity
    </div>
</div>

<div id="modalOverlay">
    <div id="modalContent">
        <div class="modal-header">
            <div>
                <span id="modalTitle" style="color:var(--color-gold); font-weight:900; font-size:2.0rem;">👤 </span>
                <span id="modalSubtitle" style="color:#A0AEC0; font-size:1.1rem; margin-left:15px;"></span>
            </div>
            <button class="modal-close" onclick="closeModal()">×</button>
        </div>
        <div class="modal-body">
            <div class="modal-summary-grid">
                <div class="metric-box">
                    <h3 style="color:#A0AEC0; margin:0 0 10px 0; font-size:1.2rem;">📋 계약건수</h3>
                    <h2 id="modalCount" style="color:#FFF; margin:0; font-size: 2.8rem;">0건</h2>
                </div>
                <div class="metric-box">
                    <h3 id="modalSumLabel" style="color:#A0AEC0; margin:0 0 10px 0; font-size:1.2rem;">💰 수수료 합계</h3>
                    <h2 id="modalSum" style="color:var(--color-gold); margin:0; font-size: 2.8rem;">0만원</h2>
                </div>
            </div>
            <div class="table-container" style="max-height: 55vh; background: rgba(0,0,0,0.2); border-radius:10px;">
                <table id="modalTable">
                    <thead><tr id="modalThead"></tr></thead>
                    <tbody id="modalTbody"></tbody>
                </table>
            </div>
        </div>
    </div>
</div>

<script>
    function checkPass() {{
        const input = document.getElementById('passInput').value;
        if(input === '8984' || input === '1057') {{
            document.getElementById('passwordGate').style.display = 'none';
            const dash = document.getElementById('mainDashboard');
            dash.style.display = 'block';
            setTimeout(() => dash.style.opacity = '1', 10);
        }} else {{
            alert('비밀번호가 일치하지 않습니다.');
        }}
    }}
    document.getElementById('passInput').addEventListener('keyup', e => {{
        if(e.key === 'Enter') checkPass();
    }});

    const rawData = {json_data};
    
    const SCOPE_CONDITIONS = {{
        'ach_life_h':  (r) => String(r['보험군']).includes('생명') && String(r['계약일자']).startsWith('202601'),
        'ach_life_p':  (r) => String(r['보험군']).includes('생명') && String(r['계약일자']).startsWith('202601'),
        'ach_dmg_new': (r) => String(r['보험군']).includes('손해') && String(r['계약일자']).startsWith('202601') && String(r['지급구분']).includes('신계약'),
        'ach_dmg_gen': (r) => String(r['보험군']).includes('손해') && String(r['계약일자']).startsWith('202601') && String(r['지급구분']).includes('일반'),
        'ach_dmg_car': (r) => String(r['보험군']).includes('손해') && String(r['계약일자']).startsWith('202601') && String(r['지급구분']).includes('자동차'),
        
        'fee_life_new':(r) => String(r['보험군']).includes('생명') && String(r['지급구분']).includes('신계약'),
        'fee_life_ret':(r) => String(r['보험군']).includes('생명') && !String(r['지급구분']).includes('신계약'),
        'fee_dmg_new': (r) => String(r['보험군']).includes('손해') && String(r['지급구분']).includes('신계약'),
        'fee_dmg_ret': (r) => String(r['보험군']).includes('손해') && String(r['지급구분']).includes('유지'),
        'fee_dmg_etc': (r) => String(r['보험군']).includes('손해') && (String(r['지급구분']).match(/일반|자동차|환수/)),
        'fee_dmg_gen': (r) => String(r['보험군']).includes('손해') && String(r['지급구분']).includes('일반')
    }};

    const SCOPE_LABELS = {{
        'ach_life_h': '🔵 생명 신계약(환산)', 'ach_life_p': '🔵 생명 신계약(보험료)', 'ach_dmg_new': '🟠 손해 신계약', 'ach_dmg_gen': '⚪ 손해 일반', 'ach_dmg_car': '⚪ 손해 자동차',
        'fee_life_new': '🔵 생명 익월 수수료', 'fee_life_ret': '🟢 생명 유지', 'fee_dmg_new': '🟠 손해 익월', 'fee_dmg_ret': '🟣 손해 유지', 'fee_dmg_etc': '⚪ 손해 기타', 'fee_dmg_gen': '⚪ 손해 일반'
    }};

    function formatNumber(num) {{ return Math.round(num).toString().replace(/\\B(?=(\\d{{3}})+(?!\\d))/g, ","); }}

    function openCompModal(scope) {{ openDetailModal(scope, null, null); }}

    function openDetailModal(scope, comp, fc) {{
        let filtered = rawData.filter(SCOPE_CONDITIONS[scope]);
        if(comp && comp !== 'null') filtered = filtered.filter(r => r['제휴사명'] && String(r['제휴사명']).includes(comp));
        if(fc && fc !== 'null' && fc !== 'dummy')   filtered = filtered.filter(r => String(r['FC명']).trim() === String(fc).trim());

        const lblScope = SCOPE_LABELS[scope] || "";
        let mainTitle = fc ? `👤 ${{fc}}` : (comp ? `🏢 ${{comp}}` : `📊 전체 상세`);
        document.getElementById('modalTitle').innerText = mainTitle;
        document.getElementById('modalSubtitle').innerText = `— ${{lblScope}} 부문`;

        const isLife = lblScope.includes('생명');
        const isAch = scope.startsWith('ach_');
        
        const sumLabel = document.getElementById('modalSumLabel');
        if (isAch) {{
            sumLabel.innerText = scope.includes('_h') ? '📋 환산 합계' : '📋 보험료 합계';
        }} else {{
            sumLabel.innerText = '💰 수수료 합계';
        }}

        let cols = ['계약일자', '제휴사명', 'FC명', '지급구분', '상품군', '상품명', '계약자'];
        if(isLife) cols.push('환산성적', '보험료');
        else cols.push('보험료');
        if(!isAch) cols.push('지사수수료');

        const headers = cols.map(c => {{
            if(c === '환산성적') return '환산';
            if(c === '보험료') return isLife ? '보험료' : '손해_보험료';
            if(c === '지사수수료') return '수수료(원)';
            return c;
        }});

        document.getElementById('modalThead').innerHTML = headers.map(h => `<th>${{h}}</th>`).join('');

        let tbodyHtml = '';
        let sumMain = 0;
        let refundSum = 0;

        filtered.forEach(r => {{
            const valForSum = isAch ? (scope.includes('_h') ? (r['환산성적']||0) : (r['보험료']||0)) : (r['지사수수료']||0);
            sumMain += valForSum;
            if(!isAch && valForSum < 0) refundSum += valForSum;
            
            tbodyHtml += '<tr>';
            cols.forEach(c => {{
                let val = r[c] || (typeof r[c] === 'number' ? 0 : '-');
                let align = 'center', color = '';
                if(typeof val === 'number') {{
                    align = 'right';
                    if(val < 0) color = 'color:#E63946;';
                    val = formatNumber(val);
                }}
                tbodyHtml += `<td style="text-align:${{align}}; ${{color}}">${{val}}</td>`;
            }});
            tbodyHtml += '</tr>';
        }});
        document.getElementById('modalTbody').innerHTML = tbodyHtml;

        document.getElementById('modalCount').innerText = `${{filtered.length}}건`;
        
        if (isAch) {{
            document.getElementById('modalSum').innerText = `${{formatNumber(sumMain)}}원`;
        }} else {{
            document.getElementById('modalSum').innerText = `${{formatNumber(sumMain / 10000)}}만원`;
        }}

        // 환수 섹션 추가 (마이너스 수수료가 있는 경우에만 요약 박스 동적 생성/업데이트)
        const summaryGrid = document.querySelector('.modal-summary-grid');
        let refundBox = document.getElementById('modalRefundBox');
        
        if (refundSum < 0) {{
            if (!refundBox) {{
                refundBox = document.createElement('div');
                refundBox.id = 'modalRefundBox';
                refundBox.className = 'metric-box';
                refundBox.style.border = '1px solid #E63946';
                summaryGrid.appendChild(refundBox);
            }}
            refundBox.innerHTML = `<h3 style="color:#E63946; margin:0 0 10px 0; font-size:1.2rem;">📉 환수(마이너스) 합계</h3>
                                 <h2 style="color:#E63946; margin:0; font-size: 2.8rem;">${{formatNumber(Math.abs(refundSum))}}원</h2>`;
            refundBox.style.display = 'block';
        }} else if (refundBox) {{
            refundBox.style.display = 'none';
        }}

        document.getElementById('modalOverlay').style.display = 'flex';
    }}

    function closeModal() {{
        document.getElementById('modalOverlay').style.display = 'none';
    }}

    document.addEventListener('keydown', e => {{
        if(e.key === 'Escape') closeModal();
    }});
</script>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print("index.html file has been successfully generated!")
