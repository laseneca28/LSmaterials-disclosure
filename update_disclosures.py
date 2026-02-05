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

# 숫자로 변환할 열
numeric_cols = ["확정 계약금액", "조건부 계약금액", "계약금액 총액(원)", "최근 매출액(원)", "매출액 대비(%)"]

def clean_number(text):
    """문자열에서 숫자와 소수점만 남기고 숫자로 변환 (예: '137,364,591,159 (원)' -> 137364591159)"""
    if not text or text == "-" or str(text).strip() == "": return 0
    # 숫자와 소수점(.)만 남기기
    cleaned = re.sub(r'[^0-9.]', '', str(text))
    if not cleaned: return 0
    try:
        # 소수점이 있으면 float, 없으면 int
        return float(cleaned) if '.' in cleaned else int(cleaned)
    except:
        return 0

def get_detailed_info(rcept_no):
    info = {c: "-" for c in cols}
    try:
        document = dart.document(rcept_no)
        tables = pd.read_html(document)
        
        for df in tables:
            # 텍스트 전처리: 공백 제거 및 문자열화
            df = df.fillna("-").astype(str)
            for _, row in df.iterrows():
                row_list = [c.replace(" ", "").replace("\n", "") for c in row.tolist()]
                
                for idx, cell in enumerate(row_list):
                    # 값 후보 찾기: 현재 칸 이후에 '-'가 아닌 첫 번째 칸을 선택
                    target_val = "-"
                    if idx + 1 < len(row_list):
                        for next_cell in row_list[idx+1:]:
                            if next_cell != "-" and next_cell != "":
                                target_val = next_cell
                                break

                    # 키워드 매칭 (더 포괄적으로 변경)
                    if "판매" in cell and "공급" in cell and "내용" in cell: info["판매ㆍ공급계약 내용"] = target_val
                    elif "조건부" in cell and "여부" in cell: info["조건부 계약여부"] = target_val
                    elif "확정" in cell and "계약금액" in cell: info["확정 계약금액"] = target_val
                    elif "조건부" in cell and "계약금액" in cell: info["조건부 계약금액"] = target_val
                    elif "계약금액" in cell and "총액" in cell: info["계약금액 총액(원)"] = target_val
                    elif "최근" in cell and "매출액" in cell: info["최근 매출액(원)"] = target_val
                    elif "매출액" in cell and "대비" in cell: info["매출액 대비(%)"] = target_val
                    elif "계약상대방" in cell: info["계약 상대방"] = target_val
                    elif "시작일" in cell: info["시작일"] = target_val
                    elif "종료일" in cell: info["종료일"] = target_val
                    elif "계약" in cell and "일자" in cell and "수주" in cell: info["계약(수주)일자"] = target_val
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
        
        # 숫자 컬럼 변환 및 엑셀 서식 적용
        for col in numeric_cols:
            if col in final_df.columns:
                final_df[col] = final_df[col].apply(clean_number)

        final_df.to_excel(writer, sheet_name=company['name'], index=False)

print("작업 완료.")
