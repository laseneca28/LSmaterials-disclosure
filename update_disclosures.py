import OpenDartReader
import pandas as pd
import os
from datetime import datetime

# 1. 설정
API_KEY = os.environ.get('DART_API_KEY') 
dart = OpenDartReader(API_KEY)
company_code = "417200" # LS머트리얼즈

def get_detailed_info(rcept_no):
    """공시 본문에서 11가지 요청 항목을 추출"""
    # 기본값 설정
    info = {
        "판매ㆍ공급계약 내용": "-", "조건부 계약여부": "-", "확정 계약금액": "-",
        "조건부 계약금액": "-", "계약금액 총액(원)": "-", "최근 매출액(원)": "-",
        "매출액 대비(%)": "-", "계약 상대방": "-", "시작일": "-", "종료일": "-", "계약(수주)일자": "-"
    }
    
    try:
        document = dart.document(rcept_no)
        tables = pd.read_html(document)
        
        for df in tables:
            # 데이터프레임을 문자열 리스트로 변환하여 탐색
            df_str = df.astype(str)
            for i, row in df_str.iterrows():
                row_text = "".join(row.values)
                
                # 항목별 매칭 (DART 표준 서식 기준)
                if "판매ㆍ공급계약 내용" in row_text: info["판매ㆍ공급계약 내용"] = row.iloc[-1]
                if "조건부 계약여부" in row_text: info["조건부 계약여부"] = row.iloc[-1]
                if "확정 계약금액" in row_text: info["확정 계약금액"] = row.iloc[-1]
                if "조건부 계약금액" in row_text: info["조건부 계약금액"] = row.iloc[-1]
                if "계약금액 총액" in row_text: info["계약금액 총액(원)"] = row.iloc[-1]
                if "최근 매출액" in row_text: info["최근 매출액(원)"] = row.iloc[-1]
                if "매출액 대비" in row_text: info["매출액 대비(%)"] = row.iloc[-1]
                if "계약 상대방" in row_text: info["계약 상대방"] = row.iloc[-1]
                if "계약기간" in row_text or "시작일" in row_text: 
                    # 보통 시작일과 종료일은 한 줄이나 인접한 줄에 있음
                    info["시작일"] = row.iloc[-1] if "시작일" in row_text else info["시작일"]
                if "종료일" in row_text: info["종료일"] = row.iloc[-1]
                if "계약(수주)일자" in row_text: info["계약(수주)일자"] = row.iloc[-1]
        
        return info
    except:
        return info

# 2. 공시 리스트 수집
df_list = dart.list(company_code, start='2024-01-01')

if df_list is not None and not df_list.empty:
    # 수주 공시만 필터링
    contracts_list = df_list[df_list['report_nm'].str.contains('단일판매|공급계약', na=False)].copy()
    
    detailed_data = []
    for idx, row in contracts_list.iterrows():
        print(f"추출 중: {row['report_nm']} ({row['rcept_no']})")
        details = get_detailed_info(row['rcept_no'])
        # 기본 정보(공시명, 날짜 등)와 상세 정보를 합침
        details['공시제목'] = row['report_nm']
        details['공시번호'] = row['rcept_no']
        detailed_data.append(details)

    # 3. 엑셀 저장용 데이터프레임 생성 및 열 순서 정렬
    final_df = pd.DataFrame(detailed_data)
    columns_order = [
        "공시제목", "판매ㆍ공급계약 내용", "조건부 계약여부", "확정 계약금액", 
        "조건부 계약금액", "계약금액 총액(원)", "최근 매출액(원)", "매출액 대비(%)", 
        "계약 상대방", "시작일", "종료일", "계약(수주)일자", "공시번호"
    ]
    # 존재하는 컬럼만 선별하여 정렬
    final_df = final_df[[col for col in columns_order if col in final_df.columns]]

    # 4. 저장
    file_name = "LS_Materials_Contracts_Detailed.xlsx"
    final_df.to_excel(file_name, index=False)
    print(f"저장 완료: {file_name}")
else:
    print("조회된 공시가 없습니다.")
