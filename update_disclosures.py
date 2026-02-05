import OpenDartReader
import pandas as pd
import os

# 1. 설정
API_KEY = os.environ.get('DART_API_KEY') 
dart = OpenDartReader(API_KEY)

# 대상 기업 리스트
target_companies = [
    {"name": "LS머트리얼즈", "code": "417200"},
    {"name": "비나텍", "code": "126340"}
]

# 11개 항목 정의
cols = ["공시제목", "판매ㆍ공급계약 내용", "조건부 계약여부", "확정 계약금액", "조건부 계약금액", 
        "계약금액 총액(원)", "최근 매출액(원)", "매출액 대비(%)", "계약 상대방", "시작일", "종료일", "계약(수주)일자"]

def get_detailed_info(rcept_no):
    info = {c: "-" for c in cols}
    try:
        document = dart.document(rcept_no)
        tables = pd.read_html(document)
        for df in tables:
            df_str = df.astype(str)
            for _, row in df_str.iterrows():
                row_text = "".join(row.values)
                if "판매ㆍ공급계약 내용" in row_text: info["판매ㆍ공급계약 내용"] = row.iloc[-1]
                if "조건부 계약여부" in row_text: info["조건부 계약여부"] = row.iloc[-1]
                if "확정 계약금액" in row_text: info["확정 계약금액"] = row.iloc[-1]
                if "조건부 계약금액" in row_text: info["조건부 계약금액"] = row.iloc[-1]
                if "계약금액 총액" in row_text: info["계약금액 총액(원)"] = row.iloc[-1]
                if "최근 매출액" in row_text: info["최근 매출액(원)"] = row.iloc[-1]
                if "매출액 대비" in row_text: info["매출액 대비(%)"] = row.iloc[-1]
                if "계약 상대방" in row_text: info["계약 상대방"] = row.iloc[-1]
                if "시작일" in row_text: info["시작일"] = row.iloc[-1]
                if "종료일" in row_text: info["종료일"] = row.iloc[-1]
                if "계약(수주)일자" in row_text: info["계약(수주)일자"] = row.iloc[-1]
        return info
    except:
        return info

# 2. 엑셀 파일 하나에 여러 시트로 저장하기
file_name = "Integrated_Disclosure_Report.xlsx"

with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
    for company in target_companies:
        print(f"[{company['name']}] 데이터 수집 중...")
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

        # 데이터가 없더라도 빈 시트 구조 생성
        final_df = pd.DataFrame(detailed_data, columns=cols)
        # 시트 이름을 회사명으로 설정하여 저장
        final_df.to_excel(writer, sheet_name=company['name'], index=False)
        print(f"[{company['name']}] 시트 생성 완료.")

print(f"최종 파일 생성 완료: {file_name}")
