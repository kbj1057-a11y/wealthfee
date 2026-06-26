import pandas as pd
import json
import os
import argparse
import sys

# 한글 출력 인코딩 설정
sys.stdout.reconfigure(encoding='utf-8')

BASE_ROOT = r'g:\내 드라이브\안티그래비티\급여,수수료PROJECT'

def export_promo_to_js():
    parser = argparse.ArgumentParser()
    parser.add_argument('yymm', help='정산년월 4자리 (예: 2606)')
    args = parser.parse_args()
    yymm = args.yymm

    INPUT_FILE = os.path.join(BASE_ROOT, '5일시책자료', f'{yymm}지급시책.xlsx')
    OUTPUT_JS = os.path.join(BASE_ROOT, 'data.js')

    if not os.path.exists(INPUT_FILE):
        print(f"Error: {INPUT_FILE} not found")
        return

    # 엑셀 로드
    df = pd.read_excel(INPUT_FILE)
    
    # 데이터 정제: FC명 결측치 제거
    df = df.dropna(subset=['FC명'])
    # 합계나 계 행 제거
    df = df[~df['FC명'].astype(str).str.contains('계|합계|소계|총계')]

    promo_records = []
    year_prefix = f"20{yymm[:2]}"
    month_suffix = yymm[2:]
    ym_num = int(f"20{yymm}") # 예: 202606
    pay_date = f"20{yymm[:2]}{yymm[2:]}05" # 예: "20260605"

    # 매핑할 항목 정의: (엑셀컬럼, 시책종류_양수, 시책종류_음수, 상품명_접미사, 카테고리_양수, 카테고리_음수)
    items_to_map = [
        ('초회', 'FC_초회', 'FC_초회환수', '초회', 'initial_cash', 'initial_refund'),
        ('2차년', 'FC_2차년', 'FC_2차년환수', '2차년', 'year2_cash', 'year2_refund'),
        ('기타', 'FC_기타', 'FC_기타환수', '기타', 'etc', 'etc'),
        ('지사전략', 'FC_전략선지급', 'FC_전략선지급환수', '지사전략', 'initial_cash', 'initial_refund'),
        ('환수', 'FC_2차년환수', 'FC_2차년환수', '환수', 'year2_refund', 'year2_refund')
    ]

    for idx, row in df.iterrows():
        fc_name = str(row['FC명']).strip()
        
        for col_name, kind_pos, kind_neg, name_suffix, cat_pos, cat_neg in items_to_map:
            if col_name not in df.columns:
                continue
                
            val = row[col_name]
            # NaN 체크 및 0 체크
            if pd.isna(val) or val == 0:
                continue
            
            try:
                val = int(round(float(val)))
            except:
                continue
                
            if val == 0:
                continue

            # 시책종류, 상품명, 카테고리 정의
            if val > 0:
                kind = kind_pos
                prod_name = f"{int(month_suffix)}월 지급시책 {name_suffix}"
                category = cat_pos
            else:
                kind = kind_neg
                prod_name = f"{int(month_suffix)}월 지급시책 {name_suffix} 환수"
                category = cat_neg

            record = {
                "정산년월": ym_num,
                "보험군": "현금",
                "제휴사명": "지사공통",
                "시책종류": kind,
                "FC명": fc_name,
                "증권번호": "-",
                "상품명": prod_name,
                "계약일자": pay_date,
                "보험료": 0,
                "시책금": val,
                "시책상세내용": f"{yymm[:2]}년 {month_suffix}월 지급시책 {name_suffix}",
                "카테고리": category
            }
            promo_records.append(record)

    key = f"20{yymm[:2]}_{yymm[2:]}"
    print(f"Parsed {len(promo_records)} promo rows from {INPUT_FILE}")

    # 기존 data.js 로드하여 보존
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

    # 신규 시책 데이터 주입
    final_data["promo_data"][key] = promo_records

    # data.js 쓰기
    with open(OUTPUT_JS, 'w', encoding='utf-8') as f:
        f.write(f"const rawData = {json.dumps(final_data, ensure_ascii=False, indent=2)};")
    
    print(f"Successfully merged & exported {len(promo_records)} promo rows to data.js for key {key}")

if __name__ == '__main__':
    export_promo_to_js()
