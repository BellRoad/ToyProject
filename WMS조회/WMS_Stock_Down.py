import requests
import pandas as pd
from io import StringIO
import winreg
import time
import os

# 날짜 변수 예시 (실제 값으로 대체)
datefrom = "20250401"
dateto = "20250430"

# 프로그램 시작 시간 기록
start_time = time.time()

# compcd별로 데이터프레임을 저장할 딕셔너리
compcd_dict = {
    "1": [],
    "2": [],
    "3": []
}

# 세션 생성 (로그인 필요시 여기에 추가)
session = requests.Session()

url1 = "http://order.mydongsim.com/Login.do?user_id=finance3&password=6093"
response1 = session.get(url1)

for i in range(1, 7):
    if i == 1:
        compcd = "1" #동심
        agency = "9999"
        agy_name = "안성"
    elif i == 2:
        compcd = "1" #동심
        agency = "3333"
        agy_name = "교원"
    elif i == 3:
        compcd = "2" #5월5일
        agency = "9999"
        agy_name = "안성"
    elif i == 4:
        compcd = "2" #5월5일
        agency = "3333"
        agy_name = "교원"
    elif i == 5:
        compcd = "3" #킨더
        agency = "9999"
        agy_name = "안성"
    elif i == 6:
        compcd = "3" #킨더
        agency = "3333"
        agy_name = "교원"

    down_url = (
        "http://order.mydongsim.com/Report/Report090excel.jsp?"
        f"comp_cd={compcd}&date_from={datefrom}&date_to={dateto}"
        f"&agy_name={agy_name}&agency={agency}"
        "&goods_name=&goods=&rule_cd1=&rule_cd2=&rule_cd3=&badgb=1"
    )

    response = session.get(down_url)
    response.encoding = 'utf-8'  # 필요시 인코딩 지정

    # HTML 테이블 읽기
    dfs = pd.read_html(StringIO(response.text))
    df = dfs[0]

    # 첫 번째 행을 헤더로 설정
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)

    # 1. '품목코드', '품목명'을 제외한 나머지 컬럼 숫자형 변환
    for col in df.columns:
        if col not in ['품목코드', '품목명']:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    # compcd별로 데이터프레임 저장
    compcd_dict[compcd].append(df)

# 다운로드 폴더 경로 가져오기
def get_download_folder():
    try:
        # 레지스트리에서 다운로드 폴더 경로 확인
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
            downloads_path = winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
        return downloads_path
    except Exception:
        # 레지스트리에서 찾지 못한 경우 기본 경로 반환
        return os.path.join(os.path.expanduser('~'), 'Downloads')

downloads_folder = get_download_folder()

# compcd별 한글명 매핑
compcd_name_map = {
    "1": "동심",
    "2": "5월5일",
    "3": "킨더"
}

# compcd별로 합쳐서 후처리
for compcd, df_list in compcd_dict.items():
    if df_list:
        # 2. 합치기, 품목명순 정렬, 품목명별 합계
        df_all = pd.concat(df_list, ignore_index=True)
        # 품목코드, 품목명으로 그룹화하여 합계
        group_cols = ['품목코드', '품목명']
        agg_cols = [col for col in df_all.columns if col not in group_cols]
        df_all = df_all.groupby(group_cols, as_index=False)[agg_cols].sum()
        # 품목명순 정렬
        df_all = df_all.sort_values('품목명').reset_index(drop=True)

        # 3. '출고 합계' 컬럼 추가
        for col in ['출고 수량', '반품 수량', 'A/S 수량', '조정 수량', '가공 수량']:
            if col not in df_all.columns:
                df_all[col] = 0  # 혹시 없는 컬럼이 있으면 0으로 추가
        df_all['출고 합계'] = (
            df_all['출고 수량'] +
            df_all['반품 수량'] +
            df_all['A/S 수량'] +
            df_all['조정 수량'] +
            df_all['가공 수량']
        )

        # 4. 컬럼 순서 지정
        col_order = [
            '품목코드', '품목명', '전재고 수량', '전재고 금액', '입고 수량',
            '출고 합계', '현재고 수량', '현재고 금액',
            '출고 수량', '반품 수량', 'A/S 수량', '조정 수량', '가공 수량',
            '입고 금액', '출고 금액', '반품 금액', 'A/S 금액', '조정 금액', '가공 금액'
        ]
        # 없는 컬럼은 자동으로 제외
        col_order = [col for col in col_order if col in df_all.columns]
        df_all = df_all[col_order]

        # 파일명 한글 처리
        kor_name = compcd_name_map.get(compcd, compcd)
        filename = f"재고리스트(누계) {kor_name} {datefrom}-{dateto}.xlsx"
        filepath = os.path.join(downloads_folder, filename)
        df_all.to_excel(filepath, index=False)
        print(f"{filename} 저장 완료")
    else:
        print(f"compcd={compcd} 데이터 없음")

# 프로그램 종료 시간 기록
end_time = time.time()

# 전체 실행 시간 계산 및 출력
elapsed_time = end_time - start_time
print(f"프로그램 실행 시간: {elapsed_time:.2f}초")

os.system("pause")