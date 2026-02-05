import OpenDartReader
import pandas as pd
from datetime import datetime
import os

# 1. 설정
API_KEY = os.environ.get('DART_API_KEY') 
dart = OpenDartReader(API_KEY)
company_code = "417200" # LS머트리얼즈

def get_detailed_info(rcept_no):
    """공시 번호로 상세 계약 정보(금액, 상대방)를 추출"""
    try:
        # 공시 본문을 HTML로 가져온 뒤 테이블(표)만 추출합니다.
        document = dart.document(rcept_no)
        tables = pd.read_html(document)
        
        amount = "-"
        partner = "-"
        
        for df in tables:
            # 표 내용을 문자열로 합쳐서 키워드 검색
            full_text = df.astype(str).values.flatten()
            text_blob = "".join(full_text)
            
            if '계약금액' in text_blob or '계약상대방' in text_blob:
                # 표 안에서 해당 항목의 위치를 찾아 값을 가져옵니다.
                for i, row in df.iterrows():
                    row_list = row.astype(str).tolist()
                    for idx, cell in enumerate(row_list):
                        if '계약금액' in cell and idx + 1 < len(row_list):
                            amount = row_list[idx + 1]
                        if '계약상대방' in cell and idx + 1 < len(row_list):
                            partner = row_list[idx + 1]
                return amount, partner
    except Exception as e:
        print(f"Error parsing {rcept_no}: {e}")
        return "-", "-"
    return "-", "-"

# 2. 최근 공시 리스트 수집 (bgn_de -> start로 수정)
# '2024-01-01' 형식으로 입력해야 합니다.
df_list = dart.list(company_code, start='2024-01-01')

if df_list is not None and not df_list.empty:
    # 3. '수주 공시'만 필터링
    contracts = df_list[df_list['report_nm'].str.contains('단일판매|공급계약', na=False)].copy()

    # 4. 상세 정보 추가
    amounts = []
    partners = []

    for idx, row in contracts.iterrows():
        amt, pt = get_detailed_info(row['rcept_no'])
        amounts.append(amt)
        partners.append(pt)

    contracts['계약금액'] = amounts
    contracts['계약상대방'] = partners

    # 5. 저장
    file_name = "LS_Materials_Contracts.xlsx"
    contracts.to_excel(file_name, index=False)
    print(f"성공! {len(contracts)}건의 데이터를 저장했습니다.")
else:
    print("공시 리스트가 비어있습니다.")
