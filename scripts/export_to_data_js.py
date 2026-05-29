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

    COL_HWANSAN_DASH = '\ud658\uc0b0\uc131\uc801' # 환산성적
    COL_FEE_DASH = '\ucd5c\uc885\uc218\uc218\ub8cc' # 최종수수료
    COL_GROUP = '\ubcf4\ud5d8\uad70'
    
    VAL_LIFE = '\uc0dd\uba85'
    VAL_NONLIFE = '\uc190\ud574'

    xl = pd.ExcelFile(INPUT_FILE)
    all_rows = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(INPUT_FILE, sheet_name=sheet)
        
        # English Sheet -> Korean Group
        is_life_sheet = 'Life' in sheet and 'NonLife' not in sheet
        group_val = VAL_LIFE if is_life_sheet else VAL_NONLIFE
        
        # 시트별 성능 컬럼을 '환산성적'으로 통일 변환
        col_hwansan_src = '\ud658\uc0b0' if is_life_sheet else '\uc218\uc815\ubcf4\ud5d8\ub8cc' # 환산 / 수정보험료
        
        if col_hwansan_src in df.columns:
            if COL_HWANSAN_DASH in df.columns and col_hwansan_src != COL_HWANSAN_DASH:
                df.drop(columns=[COL_HWANSAN_DASH], inplace=True)
            df.rename(columns={col_hwansan_src: COL_HWANSAN_DASH}, inplace=True)
            
        if COL_FEE_DASH in df.columns:
            # 최종수수료는 이름 그대로 유지
            pass
        
        if COL_GROUP not in df.columns or df[COL_GROUP].isnull().all():
            df[COL_GROUP] = group_val
            
        all_rows.extend(df.to_dict(orient='records'))

    key = f"20{yymm[:2]}_{yymm[2:]}"
    
    # 기존 data.js 파일을 읽어 기존의 월 데이터를 복구하여 유지
    final_data = {"fee_data": {}, "promo_data": {}}
    if os.path.exists(OUTPUT_JS):
        try:
            with open(OUTPUT_JS, 'r', encoding='utf-8') as f:
                content = f.read().strip()
                if content.startswith("const rawData ="):
                    json_str = content[len("const rawData ="):].strip()
                    if json_str.endswith(";"):
                        json_str = json_str[:-1].strip()
                    existing_data = json.loads(json_str)
                    if isinstance(existing_data, dict):
                        if "fee_data" in existing_data:
                            final_data["fee_data"].update(existing_data["fee_data"])
                        if "promo_data" in existing_data:
                            final_data["promo_data"].update(existing_data["promo_data"])
            print(f"Loaded existing months from data.js: Fee={list(final_data['fee_data'].keys())}, Promo={list(final_data['promo_data'].keys())}")
        except Exception as e:
            print(f"Warning: Could not parse existing data.js: {e}")
            
    # 새 데이터를 덮어쓰거나 새로 할당
    final_data["fee_data"][key] = all_rows
    
    with open(OUTPUT_JS, 'w', encoding='utf-8') as f:
        f.write(f"const rawData = {json.dumps(final_data, ensure_ascii=False, indent=2)};")
    print(f"Successfully merged & exported {len(all_rows)} rows to data.js for key {key}")

if __name__ == '__main__':
    export_to_js()
