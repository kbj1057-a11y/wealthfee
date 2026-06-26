import pandas as pd
import os
import argparse

# 기본 경로 설정
BASE_ROOT = r'g:\내 드라이브\안티그래비티\급여,수수료PROJECT'

# 영문 키 사용 (부식 방지 및 안정성) -> verified.xlsx가 한글화되었으므로 이에 맞게 수정
COL_FINAL_FEE = '\ucd5c\uc885\uc218\uc218\ub8cc' # 최종수수료
COL_FEE_FC = 'FC\uc218\uc218\ub8cc'
COL_FEE_BRANCH = '\uc9c0\uc0ac\uc218\uc218\ub8cc' 
COL_FEE_SHARE = '\ubd84\ub2f4\uae08'

def optimize_data():
    parser = argparse.ArgumentParser()
    parser.add_argument('yymm', nargs='?', default='2604')
    args = parser.parse_args()
    YYMM = args.yymm
    INPUT_FILE = os.path.join(BASE_ROOT, 'data', f'master_commission_{YYMM}_verified.xlsx')
    OUTPUT_FILE = os.path.join(BASE_ROOT, 'data', f'master_commission_{YYMM}_optimized.xlsx')

    if not os.path.exists(INPUT_FILE): return

    xl = pd.ExcelFile(INPUT_FILE)
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        for sheet_name in xl.sheet_names:
            df = pd.read_excel(INPUT_FILE, sheet_name=sheet_name)
            
            df_filtered = df[
                (df[COL_FINAL_FEE].fillna(0) != 0) | 
                (df[COL_FEE_FC].fillna(0) != 0) | 
                (df[COL_FEE_BRANCH].fillna(0) != 0) | 
                (df[COL_FEE_SHARE].fillna(0) != 0)
            ]
            
            df_filtered.to_excel(writer, sheet_name=sheet_name, index=False)

if __name__ == '__main__':
    optimize_data()
