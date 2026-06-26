import os
import pandas as pd
import argparse
import sys

# 표준 24개 컬럼 정의 (Unicode Escapes) - 분석 결과에 따른 최적화
STANDARD_24_COLS = [
    'NO', '\uc815\uc0b0\ub144\uc6d4', '\uc81c\ud734\uc0ac\uba85', '\uc99d\uad8c\ubc88\ud638', 
    '\ubcf8\uc0ac', '\uc0ac\uc5c5\ub2e8', '\uc9c0\uc0ac', '\uc9c0\uc810', '\ud300', 
    '\uc0ac\ubc88', 'FC\uba85', '\uc9c0\uae09\uad6c\ubd84', '\uae30\ucd08\uc0c1\ud0dc', 
    '\ud45c\uc900\uc0c1\ud0dc', '\uc0c1\ud488\uad70', '\uc0c1\ud488\ucf54\ub4dc', 
    '\uc0c1\ud488\uba85', '\uacc4\uc57d\uc77c\uc790', '\ud0dc\uc544\uad6c\ubd84', 
    '\uacc4\uc57d\uc790', '\ub0a9\uc785\ubc29\ubc95', '\ub0a9\uc785\uae30\uac04', 
    '\ub0a9\uc785\ud68c\ucc28', '\ubcf4\ud5d8\ub8cc'
]

def process_files():
    parser = argparse.ArgumentParser()
    parser.add_argument('yymm', nargs='?', default='2604')
    args = parser.parse_args()
    YYMM = args.yymm
    
    BASE_ROOT = r'g:\내 드라이브\안티그래비티\급여,수수료PROJECT'
    SEARCH_ROOT = os.path.join(BASE_ROOT, '27\uc77c\uc218\uc218\ub8cc\uc790\ub8cc')
    OUTPUT_FILE = os.path.join(BASE_ROOT, 'data', f'master_commission_{YYMM}_refined.xlsx')

    life_data = []
    nonlife_data = []
    
    target_files = []
    for root, dirs, files in os.walk(SEARCH_ROOT):
        if YYMM in root:
            for f in files:
                if f.endswith('.xlsx') and '\uc774\ud30c\ud2b8\ub108' not in f and not f.startswith('~$'):
                    target_files.append(os.path.join(root, f))
    
    for f in target_files:
        try:
            xl = pd.ExcelFile(f)
            for sheet in xl.sheet_names:
                try:
                    df = pd.read_excel(xl, sheet_name=sheet)
                    if df.shape[1] < 15: continue 
                    
                    target_df = df.copy()
                    is_life = '\uc0dd\uba85' in sheet or '\uc0dd\uba85' in f
                    
                    for i in range(len(target_df)):
                        try:
                            # 24열 표준 구조로 변환 (부족하면 채움)
                            row_vals = target_df.iloc[i, :24].tolist()
                            if len(row_vals) < 24:
                                row_vals += [None] * (24 - len(row_vals))
                            
                            subset = pd.Series(row_vals, index=STANDARD_24_COLS)
                            
                            # 가변적인 컬럼 구조 대응
                            max_idx = target_df.shape[1] - 1
                            comm = target_df.iloc[i, max_idx]
                            
                            def to_f(v):
                                try: return float(str(v).replace(',', '').strip()) if not pd.isna(v) else 0
                                except: return 0
                            
                            f_comm = to_f(comm)
                            
                            if is_life:
                                perf = target_df.iloc[i, 24] if max_idx >= 24 else 0
                                f_perf = to_f(perf)
                                subset['val_hwansan'] = f_perf
                                subset['val_premium'] = 0
                            else:
                                premium_24 = target_df.iloc[i, 24] if max_idx >= 24 else 0
                                f_premium = to_f(premium_24)
                                # 24번째 열 수정보험료가 없거나 0이면 원본 23번째 열의 보험료(subset['보험료'])를 폴백으로 사용
                                subset['val_premium'] = f_premium if f_premium != 0 else to_f(subset['\ubcf4\ud5d8\ub8cc'])
                                subset['val_hwansan'] = 0
                                
                            subset['val_fee'] = f_comm
                            
                            if is_life: life_data.append(subset)
                            else: nonlife_data.append(subset)
                        except Exception: 
                            continue
                except Exception:
                    continue
        except Exception:
            continue

    print(f"Collected Life Rows: {len(life_data)}, Non-Life Rows: {len(nonlife_data)}")
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    
    # openpyxl 엔진 사용 (Multi-sheet 지원 보장)
    try:
        writer = pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl')
        
        if life_data: pd.DataFrame(life_data).to_excel(writer, sheet_name='Life', index=False)
        else: pd.DataFrame(columns=STANDARD_24_COLS).to_excel(writer, sheet_name='Life', index=False)
            
        if nonlife_data: pd.DataFrame(nonlife_data).to_excel(writer, sheet_name='NonLife', index=False)
        else: pd.DataFrame(columns=STANDARD_24_COLS).to_excel(writer, sheet_name='NonLife', index=False)
            
        writer.close()
        return True
    except Exception as e:
        print(f"Excel Save Error: {e}")
        # 폴백: 단일 시트라도 저장
        if life_data or nonlife_data:
            pd.concat((life_data + nonlife_data), ignore_index=True).to_excel(OUTPUT_FILE, index=False)
            return True
        return False
    return True

if __name__ == '__main__':
    process_files()
