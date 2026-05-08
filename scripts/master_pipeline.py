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
                              encoding='utf-8')
        if result.returncode != 0:
            print(f"Error in {step_name}:")
            print(result.stdout)
            print(result.stderr)
            return False
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

    print(f"========== Monthly Commission Pipeline START ({YYMM}) ==========")

    # Stage 1: Merge
    if not run_step("Stage 1 - Insurer Merge", ["scripts/merge_commission_data.py", YYMM]):
        return

    # Stage 2: Enrich (Matching e-Partner)
    if not run_step("Stage 2 - e-Partner Reconciliation", ["scripts/enrich_with_epartner.py", YYMM]):
        return

    # Stage 3: Optimize (Filter zero rows)
    if not run_step("Stage 3 - Optimization", ["scripts/optimize_master_file.py", YYMM]):
        return

    # Stage 4: Export to JS
    if not run_step("Stage 4 - Export to Dashboard", ["scripts/export_to_data_js.py", YYMM]):
        return

    print(f"\n========== Pipeline COMPLETED (2604) ==========")
    print(f"Final Data generated: g:\\내 드라이브\\안티그래비티\\급여,수수료PROJECT\\data.js")

if __name__ == '__main__':
    main()
