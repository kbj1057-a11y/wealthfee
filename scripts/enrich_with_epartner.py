import pandas as pd
import os
import argparse

# 기본 경로 설정
BASE_ROOT = r'g:\내 드라이브\안티그래비티\급여,수수료PROJECT'

# Unicode Escapes (부식 방지)
COL_POLICY = '\uc99d\uad8c\ubc88\ud638' # 증권번호
COL_COMPANY = '\uc81c\ud734\uc0ac\uba85' # 제휴사명
COL_FEE_FC = 'FC\uc218\uc218\ub8cc'      # FC수수료
COL_FEE_BRANCH = '\uc9c0\uc0ac\uc218\uc218\ub8cc' # 지사수수료
COL_FEE_SHARE = '\ubd84\ub2f4\uae08'     # 분담금

def normalize_policy(val):
    if pd.isna(val): return ""
    s = str(val).strip()
    if s.endswith('.0'): s = s[:-2]
    return s

def normalize_company(name):
    if not isinstance(name, str): return ""
    for suffix in ['\ubcf4\ud5d8', '\ud654\uc7ac', '\uc0dd\uba85', '\ud574\uc0c1', '\uc130\ud574']: # 보험,화재,생명,해상,손해
        name = name.replace(suffix, "")
    return name.strip()

def enrich_data():
    parser = argparse.ArgumentParser()
    parser.add_argument('yymm', nargs='?', default='2604')
    args = parser.parse_args()
    YYMM = args.yymm
    YEAR_SHORT = YYMM[:2]
    FOLDER_YEAR = f"{YEAR_SHORT}\ub144\uc218\uc218\ub8cc\uc790\ub8cc"

    FILE1_PATH = os.path.join(BASE_ROOT, 'data', f'master_commission_{YYMM}_refined.xlsx')
    FILE2_PATH = os.path.join(BASE_ROOT, '27\uc77c\uc218\uc218\ub8cc\uc790\ub8cc', FOLDER_YEAR, f'{YYMM}\uc218\uc785\uc218\uc218\ub8cc\uc0b0\ucd9c', f'{YYMM}\uc774\ud30c\ud2b8\ub108\uc218\uc218\ub8cc\uc5c5\ub370\uc774\ud2b8.xlsx')
    OUTPUT_PATH = os.path.join(BASE_ROOT, 'data', f'master_commission_{YYMM}_verified.xlsx')

    if not os.path.exists(FILE1_PATH): return

    # e-Partner 데이터 로드 (Robust reading)
    try:
        # 이파트너 파일은 시트명이 고정임
        df_mgmt = pd.read_excel(FILE2_PATH, sheet_name='\uad00\ub9ac\uc218\uc218\ub8cc') # 관리수수료
        df_life_v2 = pd.read_excel(FILE2_PATH, sheet_name='\uc0dd\uba85\ubcf4\ud5d8') # 생명보험
        df_janggi_v2 = pd.read_excel(FILE2_PATH, sheet_name='\uc7a5\uae30') # 장기
        
        for df in [df_mgmt, df_life_v2, df_janggi_v2]:
            df[COL_POLICY] = df['\uc99d\uad8c\ubc88\ud638'].apply(normalize_policy)
            df['join_key'] = df['\uc81c\ud734\uc0ac'].apply(normalize_company)
        
        agg_mgmt = df_mgmt.groupby([COL_POLICY, 'join_key'])[[COL_FEE_FC, COL_FEE_BRANCH, COL_FEE_SHARE]].sum().reset_index()
        agg_life = df_life_v2.groupby([COL_POLICY, 'join_key'])[['\uc218\uc218\ub8cc\uacc4']].sum().reset_index()
        agg_life.rename(columns={'\uc218\uc218\ub8cc\uacc4': COL_FEE_BRANCH}, inplace=True)
        agg_janggi = df_janggi_v2.groupby([COL_POLICY, 'join_key'])[['\uc218\uc218\ub8cc\uacc4']].sum().reset_index()
        agg_janggi.rename(columns={'\uc218\uc218\ub8cc\uacc4': COL_FEE_BRANCH}, inplace=True)
        
        final_agg_v2 = pd.concat([agg_mgmt, agg_life, agg_janggi], ignore_index=True).groupby([COL_POLICY, 'join_key']).sum().reset_index()
    except Exception as e:
        print(f"e-Partner Load Error: {e}")
        final_agg_v2 = pd.DataFrame(columns=[COL_POLICY, 'join_key', COL_FEE_FC, COL_FEE_BRANCH, COL_FEE_SHARE])

    # Master 파일 처리 (English Sheet Names)
    xl1 = pd.ExcelFile(FILE1_PATH)
    with pd.ExcelWriter(OUTPUT_PATH, engine='openpyxl') as writer:
        for sheet_name in xl1.sheet_names: # Life, NonLife
            df_master = pd.read_excel(xl1, sheet_name=sheet_name)
            df_master[COL_POLICY] = df_master[COL_POLICY].apply(normalize_policy)
            df_master['join_key'] = df_master[COL_COMPANY].apply(normalize_company)
            
            # Merge with suffixes to handle existing fee columns
            enriched_df = pd.merge(df_master, final_agg_v2, on=[COL_POLICY, 'join_key'], how='left', suffixes=('', '_ep'))
            
            # e-Partner 데이터가 있으면 우선 적용, 없으면 기존 값 유지
            for col in [COL_FEE_FC, COL_FEE_BRANCH, COL_FEE_SHARE]:
                ep_col = f"{col}_ep"
                if ep_col in enriched_df.columns:
                    enriched_df[col] = enriched_df[ep_col].fillna(enriched_df[col])
                    enriched_df.drop(columns=[ep_col], inplace=True)
            
            mask = enriched_df.duplicated(subset=[COL_POLICY, 'join_key'], keep='first')
            enriched_df.loc[mask, [COL_FEE_FC, COL_FEE_BRANCH, COL_FEE_SHARE]] = 0
            enriched_df.drop(columns=['join_key'], inplace=True)
            
            for col in [COL_FEE_FC, COL_FEE_BRANCH, COL_FEE_SHARE]:
                enriched_df[col] = pd.to_numeric(enriched_df[col], errors='coerce').fillna(0)
            
            enriched_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
    print(f"Successfully enriched: {OUTPUT_PATH}")

if __name__ == '__main__':
    enrich_data()
