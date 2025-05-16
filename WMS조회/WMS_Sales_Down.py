import requests
import pandas as pd
from io import StringIO
import os
import winreg
import time

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

# 작업 폴더 반영
downloads_folder = get_download_folder()
os.chdir(downloads_folder)

# 날짜 등 기초 정보 입력 및 서식 검토
def get_date_input(prompt):
    while True:
        date_str = input(prompt)
        if len(date_str) != 8:
            print("입력한 날짜 형식이 올바르지 않습니다. 8자리로 다시 입력해주세요. (예: 20230101)")
            continue
        else:
            return date_str

def get_co_input(co):
    while True:
        comp_str = input(co)
        if comp_str in ("1", "2", "3", ""):
            return comp_str
        else:
            print("1,2,3 중 입력해주세요.")
            continue

while True:
    compcd = get_co_input("Enter. 전체 | 1.동심 | 2.5월5일 | 3.킨더 : ")
    datefrom = get_date_input("시작날짜(예. 20230101) : ")
    dateto = get_date_input("종료날짜(예. 20230131) : ")
    if dateto < datefrom:
        print("종료날짜는 시작날짜보다 빠를 수 없습니다. 날짜를 다시 입력해주세요.")
    else:
        break

# 프로그램 시작 시간 기록
start_time = time.time()

# 입력데이터 정리리
def datetrans(date):
    year = date[:4]
    month = date[4:6]
    day = date[6:]
    return f"{year}.{month}.{day}"
datetrans(datefrom)
datetrans(dateto)

# 사이트 접속
session = requests.Session()

url1 = "http://order.mydongsim.com/Login.do?user_id=finance3&password=6093"
response1 = session.get(url1)

url2 = "http://order.mydongsim.com/Report/Report070excel.jsp?comp_cd=" + compcd + "&date_from=" + datefrom + "&date_to=" + dateto + "&supply_cust_nm=&supply_custno=&goods_name=&goods=&rule_cd6="
# url2 = "http://order.mydongsim.com/Report/Report070excel.jsp?comp_cd=&date_from=20250101&date_to=20250131&supply_cust_nm=&supply_custno=&goods_name=&goods=&rule_cd6="

response2 = session.get(url2)

# HTML 파일 읽기
dfs = pd.read_html(StringIO(response2.text))
df = dfs[0]

# 첫 번째 행을 헤더로 설정
df.columns = df.iloc[0]
df = df[1:].reset_index(drop=True)

# 날짜 열 변환
date_columns = ['처리일자', '주문일자']
for col in date_columns:
    df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d')

# 숫자 열 변환
numeric_columns = ['수량', '단가', '금액']
for col in numeric_columns:
    df[col] = pd.to_numeric(df[col], errors='coerce')

# 나머지 열은 문자열로 유지
string_columns = ['NO', '주문번호', '지사ID', '주문자', '인수자', '주문시간', '구분', '제품코드', '제품명', '단위']
for col in string_columns:
    df[col] = df[col].astype(str)

# 엑셀로 저장
def get_compcd_name(comp):
    return {
        "" : "전체",
        "1" : "동심",
        "2" : "5월5일",
        "3" : "킨더",
    }.get(comp, "알 수 없음")
compcd = get_compcd_name(compcd)

xlsxFile = ("매출장 " + compcd + ' ' + datefrom + '-' + dateto + '.xlsx')
df.to_excel(xlsxFile, sheet_name='매출장', index=False)

print("WMS 매출장 저장 :" + xlsxFile)

# 프로그램 종료 시간 기록
end_time = time.time()

# 전체 실행 시간 계산 및 출력
elapsed_time = end_time - start_time
print(f"프로그램 실행 시간: {elapsed_time:.2f}초")

os.system("pause")