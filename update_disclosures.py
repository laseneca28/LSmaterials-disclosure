import OpenDartReader
import pandas as pd
from datetime import datetime
import os

# 1. 설정
API_KEY = os.environ.get('DART_API_KEY') # GitHub Secrets에서 가져옴
dart = OpenDartReader(API_KEY)
company_code = "417200" # LS머트리얼즈

def get_detailed_info(rcept_no):
    """공시 번호로 상세 계약 정보(금액, 상대방)를 추출"""
    try:
        # 단일판매/공급계약체결 공시의 상세 테이블을 가져옴
        df_dict = dart.document(rcept_no)
        # 보통 첫 번째나 두 번째 테이블에 계약 정보가 들어있습니다.
        for df in df_dict:
            if '계약내역' in str(df.values) or '상대방' in str(df.values):
                # 행/열 구조가 유동적이므로 키워드로 탐색
                amount = ""
                partner = ""
                for i, row in df.iterrows():
                    row_str = str(row.values)
                    if '계약금액' in row_str:
                        amount = row.iloc[1] if len(row) > 1 else ""
                    if '계약상대' in row_str:
                        partner = row.iloc[1] if len(row) > 1 else ""
                return amount, partner
    except:
        return "-", "-"
    return "-", "-"

# 2. 최근 1년 공시 리스트 수집
df_list = dart.list(company_code, bgn_de='20240101')
contracts = df_list[df_list['report_nm'].str.contains('단일판매|공급계약', na=False)].copy()

# 3. 상세 정보 추가
amounts = []
partners = []

for idx, row in contracts.iterrows():
    amt, pt = get_detailed_info(row['rcept_no'])
    amounts.append(amt)
    partners.append(pt)

contracts['계약금액'] = amounts
contracts['계약상대방'] = partners

# 4. 저장
file_name = "LS_Materials_Contracts.xlsx"
contracts.to_excel(file_name, index=False)
print(f"Updated {file_name} with {len(contracts)} items.")
