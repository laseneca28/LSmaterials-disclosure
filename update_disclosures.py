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
    """공시 본문에서 키워드 매칭을 통해 정확한 우측 값을 추출"""
    info = {c: "-" for c in cols}
    try:
        document = dart.document(rcept_no)
        tables = pd.read_html(document)
        
        for df in tables:
            df = df.fillna("-").astype(str)
            for i, row in df.iterrows():
                row_list = row.tolist()
                for idx, cell in enumerate(row_list):
                    # 공백 제거 후 키워드 비교
                    clean_cell = cell.replace(" ", "").replace("\n", "").replace("\r", "")
                    
                    # 키워드 매칭 로직 (해당 칸에 키워드가 있으면 바로 다음 칸 idx+1 을 값으로 취함)
                    if "판매ㆍ공급계약내용" in clean_cell and idx + 1 < len(row_list):
                        info["판매ㆍ공급계약 내용"] = row_list[idx+1]
                    elif "조건부계약여부" in clean_cell and idx + 1 < len(row_list):
                        info["조건부 계약여부"] = row_list[idx+1]
                    elif "확정계약금액" in clean_cell and idx + 1 < len(row_list):
                        info["확정 계약금액"] = row_list[idx+1]
                    elif "조건부계약금액" in clean_cell and idx + 1 < len(row_list):
                        info["조건부 계약금액"] = row_list[idx+1]
                    elif "계약금액총액" in clean_cell and idx + 1 < len(row_list):
                        info["계약금액 총액(원)"] = row_list[idx+1]
                    elif "최근매출액" in clean_cell and idx + 1 < len(row_list):
                        info["최근 매출액(원)"] = row_list[idx+1]
                    elif "매출액대비" in clean_cell and idx + 1 < len(row_list):
                        info["매출액 대비(%)"] = row_list[idx+1]
                    elif "계약상대방" in clean_cell and idx + 1 < len(row_list):
                        # '3. 계약상대방' 등 숫자가 붙은 경우도 포함
                        info["계약 상대방"] = row_list[idx+1]
                    elif "시작일" in clean_cell and idx + 1 < len(row_list):
                        info["시작일"] = row_list[idx+1]
                    elif "종료일" in clean_cell and idx + 1 < len(row_list):
                        info["종료일"] = row_list[idx+1]
                    elif "계약(수주)일자" in clean_cell and idx + 1 < len(row_list):
                        info["계약(수주)일자"] = row_list[idx+1]
        
        return info
    except:
        return info

# 2. 메인 실행 로직
file_name = "Integrated_Disclosure_Report.xlsx"

with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
    for company in target_companies:
        print(f"[{company['name']}] 데이터 정밀 수집 중...")
        try:
            # 2024년부터 현재까지의 공시 리스트
            df_list = dart.list(company['code'], start='2024-01-01')
        except:
            df_list = None

        detailed_data = []
        if df_list is not None and not df_list.empty:
            # '단일판매' 또는 '공급계약' 키워드가 들어간 공시만 필터링
            contracts_list = df_list[df_list['report_nm'].str.contains('단일판매|공급계약', na=False)].copy()
            for _, row in contracts_list.iterrows():
                details = get_detailed_info(row['rcept_no'])
                details['공시제목'] = row['report_nm']
                detailed_data.append(details)

        final_df = pd.DataFrame(detailed_data, columns=cols)
        final_df.to_excel(writer, sheet_name=company['name'], index=False)

print("모든 시트 생성 완료.")
