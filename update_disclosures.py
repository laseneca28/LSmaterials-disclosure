import OpenDartReader
import pandas as pd
import os

# 1. 설정
API_KEY = os.environ.get('DART_API_KEY') 
dart = OpenDartReader(API_KEY)
company_code = "417200" # LS머트리얼즈

# 요청하신 11개 항목 정의
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

# 2. 공시 리스트 수집 (start 매개변수 사용)
try:
    df_list = dart.list(company_code, start='2024-01-01')
except:
    df_list = None

detailed_data = []

# 데이터가 있는 경우에만 상세 정보 추출
if df_list is not None and not df_list.empty:
    contracts_list = df_list[df_list['report_nm'].str.contains('단일판매|공급계약', na=False)].copy()
    for _, row in contracts_list.iterrows():
        details = get_detailed_info(row['rcept_no'])
        details['공시제목'] = row['report_nm']
        detailed_data.append(details)

# 3. 데이터프레임 생성 (데이터가 없어도 컬럼 구조는 유지)
final_df = pd.DataFrame(detailed_data, columns=cols)

# 4. 저장 (에러 방지를 위해 항상 파일 생성)
final_df.to_excel("LS_Materials_Contracts.xlsx", index=False)
print(f"작업 완료: {len(detailed_data)}건의 데이터 처리됨.")
