import json
import os

def verify_data_js():
    data_js_path = r'g:\내 드라이브\안티그래비티\급여,수수료PROJECT\data.js'
    if not os.path.exists(data_js_path):
        print("data.js not found.")
        return

    with open(data_js_path, 'r', encoding='utf-8') as f:
        content = f.read()
        
    # Extract JSON object
    try:
        obj_start = content.find('{')
        obj_end = content.rfind('}')
        data = json.loads(content[obj_start:obj_end+1])
    except Exception as e:
        print(f"Failed to parse JSON: {e}")
        return

    rows = data.get('fee_data', {}).get('2026_04', [])
    if not rows:
        print("No 2026_04 data found in data.js")
        return

    # Check for Korean keys
    sample = rows[0]
    keys = list(sample.keys())
    print(f"Sample row keys: {keys}")
    
    # Check for specific expected keys
    expected = ['제휴사명', '보험료', '지급구분', '계약일자', 'FC명']
    found = [k for k in expected if k in keys]
    missing = [k for k in expected if k not in keys]
    
    print(f"Found expected keys: {found}")
    if missing:
        print(f"CRITICAL MISSING KEYS: {missing}")
    else:
        print("SUCCESS: All essential Korean keys are present and clean.")

    # Check the sums for General/Auto
    gen_sum = sum(r.get('보험료', 0) for r in rows if '일반' in str(r.get('지급구분', '')))
    car_sum = sum(r.get('보험료', 0) for r in rows if '자동차' in str(r.get('지급구분', '')))
    print(f"General Premium Sum: {gen_sum:,.0f}")
    print(f"Auto Premium Sum: {car_sum:,.0f}")

if __name__ == '__main__':
    verify_data_js()
