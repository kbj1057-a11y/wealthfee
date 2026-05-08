---
description: 월간 수입수수료 데이터를 자동으로 통합하고 검증합니다 (병합 -> 이파트너 매칭 -> 최적화)
---

수입수수료 정산 자동화 워크플로우를 시작합니다.

1. 정산 대상 년월(YYMM)을 입력받아 폴더 경로를 설정합니다.
// turbo
2. 파이썬 마스터 파이프라인을 실행합니다.
   `python scripts/master_pipeline.py 2604`

3. 최종 생성된 마스터 파일을 확인하고 합계 수치를 보고합니다.
   - [master_commission_2604_verified.xlsx](file:///g:/%EB%82%B4%20%EB%93%9C%EB%9D%BC%EC%9D%B4%EB%B8%8C/%EC%95%88%ED%8B%B0%EA%B7%B8%EB%9E%98%EB%B9%84%ED%8B%B0/%EA%B8%89%EC%97%AC,%EC%88%98%EC%88%98%EB%A3%8CPROJECT/data/master_commission_2604_verified.xlsx)

4. 작업 완료 후 최종 보고서(Walkthrough)를 업데이트합니다.
