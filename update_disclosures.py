import OpenDartReader
import pandas as pd
import os
import re

# 1. 설정
API_KEY = os.environ.get('DART_API_KEY') 
dart = OpenDartReader(API_KEY)

target_companies = [
    {"name": "LS머트리얼즈", "code": "417200"},
    {"name": "비나텍", "code": "126340"}
]

cols = ["공시제목", "판매ㆍ공급계약 내용", "조건부 계약여부", "확정 계약금액", "조건부 계약금액", 
        "계약금액 총액(원)", "최근 매출액(원)", "매출액 대비(%)", "계약 상대방", "시작일", "종료일", "계약(수주)일자"]

# 숫자로 변환할 열 리스트
numeric_cols = ["확정 계약금액", "조건부 계약금액", "계약금액 총액(원)", "최근 매출액(원)", "매출액 대비(%)"]

def clean_number(text):
    """문자열에서 숫자와 소수점만 남기고 숫자로 변환 (예: '1,234.5 (원)' -> 1234.5)"""
    if not text or text == "-": return 0
    # 숫자와 소수점(.)만 남기기
    cleaned = re.sub(r'[^0-9.]', '', str(text))
    try:
        return float(cleaned) if '.' in cleaned else int(cleaned)
    except:
        return 0

def get_detailed_info(rcept_no):
    info = {c: "-" for c in cols}
    try:
        document = dart.document(rcept_no)
        tables = pd.read_html(document)
        
        for df in tables:
            df = df.fillna("-").astype(str)
            for _, row in df.iterrows():
                row_list = [c.replace(" ", "").replace("\n", "") for c in row.tolist()]
                
                # 키워드가 포함된 행에서 데이터 추출
                for idx, cell in enumerate(row_list):
                    target_val = "-"
                    # 해당 행의 오른쪽 칸들 중 데이터가 있는 가장 먼 칸을 선택 (보통 맨 오른쪽이 값)
                    if idx + 1 < len(row_list):
                        valid_cells = [c for c in row_list[idx+1:] if c != "-" and c != ""]
                        target_val = valid_cells[-1] if valid_cells else "-"

                    if "판매ㆍ공급계약내용" in cell: info["판매ㆍ공급계약 내용"] = target_val
                    elif "조건부계약여부" in cell: info["조건부 계약여부"] = target_val
                    elif "확정계약금액" in cell: info["확정 계약금액"] = target_val
                    elif "조건부계약금액" in cell: info["조건부 계약금액"] = target_val
                    elif "계약금액총액" in cell: info["계약금액 총액(원)"] = target_val
                    elif "최근매출액" in cell: info["최근 매출액(원)"] = target_val
                    elif "매출액대비" in cell: info["매출액 대비(%)"] = target_val
                    elif "계약상대방" in cell: info["계약 상대방"] = target_val
                    elif "시작일" in cell: info["시작일"] = target_val
                    elif "종료일" in cell: info["종료일"] = target_val
                    elif "계약(수주)일자" in cell: info["계약(수주)일자"] = target_val
        return info
    except:
        return info

# 2. 실행 및 엑셀 저장
file_name = "Integrated_Disclosure_Report.xlsx"

with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
    for company in target_companies:
        print(f"[{company['name']}] 처리 중...")
        try:
            df_list = dart.list(company['code'], start='2024-01-01')
        except:
            df_list = None

        detailed_data = []
        if df_list is not None and not df_list.empty:
            contracts_list = df_list[df_list['report_nm'].str.contains('단일판매|공급계약', na=False)].copy()
            for _, row in contracts_list.iterrows():
                details = get_detailed_info(row['rcept_no'])
                details['공시제목'] = row['report_nm']
                detailed_data.append(details)

        final_df = pd.DataFrame(detailed_data, columns=cols)
        
        # 숫자 데이터 변환 작업
        for col in numeric_cols:
            if col in final_df.columns:
                final_df[col] = final_df[col].apply(clean_number)

        final_df.to_excel(writer, sheet_name=company['name'], index=False)

print("작업이 완료되었습니다.")
