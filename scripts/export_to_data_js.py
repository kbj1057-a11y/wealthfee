import pandas as pd
import json
import os
import argparse

# 기본 경로 설정
BASE_ROOT = r'g:\내 드라이브\안티그래비티\급여,수수료PROJECT'

def export_to_js():
    parser = argparse.ArgumentParser()
    parser.add_argument('yymm', nargs='?', default='2604')
    args = parser.parse_args()
    yymm = args.yymm
    INPUT_FILE = os.path.join(BASE_ROOT, 'data', f'master_commission_{yymm}_optimized.xlsx')
    OUTPUT_JS = os.path.join(BASE_ROOT, 'data.js')

    if not os.path.exists(INPUT_FILE):
        print(f"Error: {INPUT_FILE} not found")
        return

    # Unicode Escapes
    # 영문 키 -> 국문 레이블 매핑
    COL_HWANSAN_RAW = 'val_hwansan'
    COL_HWANSAN_DASH = '\ud658\uc0b0\uc131\uc801'
    COL_PREM_RAW = 'val_premium'
    COL_PREM_DASH = '\ubcf4\ud5d8\ub8cc'
    COL_FEE_RAW = 'val_fee'
    COL_FEE_DASH = '\ucd5c\uc885\uc218\uc218\ub8cc'
    COL_GROUP = '\ubcf4\ud5d8\uad70'
    
    VAL_LIFE = '\uc0dd\uba85'
    VAL_NONLIFE = '\uc190\ud574'

    xl = pd.ExcelFile(INPUT_FILE)
    all_rows = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(INPUT_FILE, sheet_name=sheet)
        
        # English Sheet -> Korean Group
        group_val = VAL_LIFE if 'Life' in sheet and 'NonLife' not in sheet else VAL_NONLIFE
        
        # 기존 데이터에 포함된 0점짜리 컬럼 제거 후 정제된 컬럼으로 대체
        if COL_HWANSAN_RAW in df.columns:
            if COL_HWANSAN_DASH in df.columns: df.drop(columns=[COL_HWANSAN_DASH], inplace=True)
            df.rename(columns={COL_HWANSAN_RAW: COL_HWANSAN_DASH}, inplace=True)
        if COL_PREM_RAW in df.columns:
            if COL_PREM_DASH in df.columns: df.drop(columns=[COL_PREM_DASH], inplace=True)
            df.rename(columns={COL_PREM_RAW: COL_PREM_DASH}, inplace=True)
        if COL_FEE_RAW in df.columns:
            if COL_FEE_DASH in df.columns: df.drop(columns=[COL_FEE_DASH], inplace=True)
            df.rename(columns={COL_FEE_RAW: COL_FEE_DASH}, inplace=True)
        
        if COL_GROUP not in df.columns or df[COL_GROUP].isnull().all():
            df[COL_GROUP] = group_val
            
        all_rows.extend(df.to_dict(orient='records'))

    key = f"20{yymm[:2]}_{yymm[2:]}"
    final_data = {"fee_data": {key: all_rows}, "promo_data": {}}
    
    with open(OUTPUT_JS, 'w', encoding='utf-8') as f:
        f.write(f"const rawData = {json.dumps(final_data, ensure_ascii=False, indent=2)};")
    print(f"Exported {len(all_rows)} rows to data.js")

if __name__ == '__main__':
    export_to_js()
