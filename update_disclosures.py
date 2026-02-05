import OpenDartReader
import pandas as pd
import os

# 1. 설정
API_KEY = os.environ.get('DART_API_KEY') 
dart = OpenDartReader(API_KEY)
company_code = "417200" # LS머트리얼즈

def get_detailed_info(rcept_no):
    """공시 본문에서 11가지 요청 항목을 추출"""
    info = {
        "판매ㆍ공급계약 내용": "-", "조건부 계약여부": "-", "확정 계약금액": "-",
        "조건부 계약금액": "-", "계약금액 총액(원)": "-", "최근 매출액(원)": "-",
        "매출액 대비(%)": "-", "계약 상대방": "-", "시작일": "-", "종료일": "-", "계약(수주)일자": "-"
    }
    try:
        document = dart.document(rcept_no)
        tables = pd.read_html(document)
        for df in tables:
            df_str = df.astype(str)
            for i, row in df_str.iterrows():
                row_text = "".join(row.values)
                # 요청하신 11개 항목 매칭
                if "판매ㆍ공급계약 내용" in row_text: info["판매ㆍ공급계약 내용"] = row.iloc[-1]
                if "조건부 계약여부" in row_text: info["조건부 계약여부"] = row.iloc[-1]
                if "확정 계약금액" in row_text: info["확정 계약금액"] = row.iloc[-1]
                if "조건부 계약금액" in row_text: info["조건부 계약금액"] = row.iloc[-1]
                if "계약금액 총액" in row_text: info["계약금액 총액(원)"] = row.iloc[-1]
                if "최근 매출액" in row_text: info["최근 매출액(원)"] = row.iloc[-1]
                if "매출액 대비" in row_text: info["매출액 대비(%)"] = row.iloc[-1]
                if "계약 상대방" in row_text: info["계약 상대방"] = row.iloc[-1]
                if "계약기간" in row_text or "시작일" in row_text: info["시작일"] = row.iloc[-1]
                if "종료일" in row_text: info["종료일"] = row.iloc[-1]
                if "계약(수주)일자" in row_text: info["계약(수주)일자"] = row.iloc[-1]
        return info
    except:
        return info

# 2. 공시 리스트 수집 (start 매개변수 사용)
df_list = dart.list(company_code, start='2024-01-01')

if df_list is not None and not df_list.empty:
    contracts_list = df_list[df_list['report_nm'].str.contains('단일판매|공급계약', na=False)].copy()
    detailed_data = []
    for idx, row in contracts_list.iterrows():
        details = get_detailed_info(row['rcept_no'])
        details['공시제목'] = row['report_nm']
        detailed_data.append(details)

    # 3. 데이터프레임 생성 및 열 순서 정렬
    final_df = pd.DataFrame(detailed_data)
    cols = ["공시제목", "판매ㆍ공급계약 내용", "조건부 계약여부", "확정 계약금액", "조건부 계약금액", 
            "계약금액 총액(원)", "최근 매출액(원)", "매출액 대비(%)", "계약 상대방", "시작일", "종료일", "계약(수주)일자"]
    final_df = final_df[cols]

    # 4. 저장 (파일명 확인 필수)
    final_df.to_excel("LS_Materials_Contracts.xlsx", index=False)
    print("엑셀 파일 생성 완료!")
else:
    print("데이터가 없습니다.")
