import subprocess
import sys
import os
import pandas as pd
import argparse

# 기본 경로 상향 설정
BASE_ROOT = r'g:\내 드라이브\안티그래비티\급여,수수료PROJECT'

def run_step(step_name, command_args):
    print(f"\n>>> Step: {step_name}")
    try:
        result = subprocess.run([sys.executable] + command_args, 
                              cwd=BASE_ROOT, 
                              capture_output=True, 
                              text=True, 
                              errors='replace') # utf-8 디코딩 에러 방지
        if result.returncode != 0:
            print(f"Error in {step_name}:")
            print(result.stdout)
            print(result.stderr)
            return False
        
        # print stdout directly if successful so we can see the logs
        print(result.stdout)
        print(f"Success: {step_name}")
        return True
    except Exception as e:
        print(f"Failed to execute {step_name}: {e}")
        return False

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('yymm', nargs='?', default='2604', help='YearMonth Code (YYMM)')
    args = parser.parse_args()
    YYMM = args.yymm

    # 윈도우 환경에서 Excel 파일 잠금으로 인한 권한 에러(Errno 13) 방지를 위해 Excel 프로세스 자동 정리
    if os.name == 'nt':
        print("\n>>> Safety Measure: Closing active Excel processes to prevent file lock errors...")
        try:
            subprocess.run(["taskkill", "/f", "/im", "excel.exe"], capture_output=True)
            print("Success: Closed active Excel instances.")
        except Exception as e:
            print(f"Warning during Excel cleanup: {e}")

    print(f"========== Monthly Commission Pipeline START ({YYMM}) ==========")

    # Stage 1: Merge
    if not run_step("Stage 1 - Insurer Merge", ["scripts/merge_commission_data.py", YYMM]):
        return

    # Stage 2: Enrich (Matching e-Partner)
    if not run_step("Stage 2 - e-Partner Reconciliation", ["scripts/enrich_with_epartner.py", YYMM]):
        return

    # Stage 3: Exception Report (NEW)
    if not run_step("Stage 3 - Exception Reporting", ["scripts/generate_exception_report.py", YYMM]):
        return

    # Stage 4: Optimize (Filter zero rows)
    if not run_step("Stage 4 - Optimization", ["scripts/optimize_master_file.py", YYMM]):
        return
    # Stage 5: Export to JS (Dashboard integration)
    # 안전 장치: export_to_data_js 실행 전에 data.js를 git checkout하여 이전 다중 월 누적본을 안전하게 원복
    print("\n>>> Safety Measure: Restoring data.js from Git HEAD to preserve historical records...")
    try:
        git_res = subprocess.run(["git", "checkout", "data.js"], 
                                 cwd=BASE_ROOT, 
                                 capture_output=True, 
                                 text=True)
        if git_res.returncode == 0:
            print("Success: Safely checked out data.js from Git HEAD.")
        else:
            print(f"Warning during git checkout: {git_res.stderr.strip()}")
    except Exception as e:
        print(f"Warning: Failed to execute git checkout safety measure: {e}")

    if not run_step("Stage 5 - Export to Dashboard", ["scripts/export_to_data_js.py", YYMM]):
        return

    print(f"\n========== Pipeline COMPLETED ({YYMM}) ==========")
    print(f"Final Data generated in data/ folder")

if __name__ == '__main__':
    main()
