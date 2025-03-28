# selenium의 webdriver를 사용하기 위한 import. 엣지드라이버 설치여부 체크 후 자동 실행
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.webdriver import WebDriver as EdgeDriver
from webdriver_manager.microsoft import EdgeChromiumDriverManager

# 페이지 로딩을 기다리는데에 사용할 time 모듈 import
import time

# 파일 다운로드 대기 걸기
import os

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

# 입력 날짜 형식 변경
def datetrans(date):
    year = date[:4]
    month = date[4:6]
    day = date[6:]
    return f"{year}.{month}.{day}"

datetrans(datefrom)
datetrans(dateto)

service = Service(executable_path=EdgeChromiumDriverManager().install())
driver = EdgeDriver(service=service)

# 엣지 드라이버에 url 주소 넣고 실행
driver.get('http://order.mydongsim.com/Login.do?user_id=finance3&password=6093')

# 페이지가 완전히 로딩되도록 2초동안 기다림
time.sleep(2)

down_url = "http://order.mydongsim.com/Report/Report070excel.jsp?comp_cd=" + compcd + "&date_from=" + datefrom + "&date_to=" + dateto + "&supply_cust_nm=&supply_custno=&goods_name=&goods=&rule_cd6="
driver.get(down_url)

# 엣지 브라우져 다운로드 경로 확인
import winreg

def get_download_folder():
    try:
        # 레지스트리에서 다운로드 폴더 경로 확인
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
            downloads_path = winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
        return downloads_path
    except Exception:
        # 레지스트리에서 찾지 못한 경우 기본 경로 반환
        return os.path.join(os.path.expanduser('~'), 'Downloads')

# 다운로드 폴더 경로 가져오기
downloads_folder = get_download_folder()

# 확인할 파일 이름
filename = "매출장.xls"
file_path = os.path.join(downloads_folder, filename)

# 파일이 다운로드될 때까지 대기
while not os.path.exists(file_path):
    time.sleep(5)

# 파일 다운로드 완료 후 엣지드라이버 종료
driver.quit()

print("xls를 xlsx로 변환 시작")

# 파일 이름 정리
def undatetrans(date):
    # .을 제외한 문자열을 반환
    return date.replace(".", "")

undatetrans(datefrom)
undatetrans(dateto)

def get_compcd_name(comp):
    return {
        "" : "전체",
        "1" : "동심",
        "2" : "5월5일",
        "3" : "킨더",
    }.get(comp, "알 수 없음")
compcd = get_compcd_name(compcd)

import pandas as pd
import os

# 파일 이름 변경
new_filename = os.path.join(downloads_folder,filename.rsplit('.', 1)[0] + ' ' + compcd + ' ' + datefrom + '-' + dateto + '.html')
os.rename(file_path, new_filename)
htmlFile = new_filename

# HTML 파일 읽기
df = pd.read_html(htmlFile)[0]  # 첫 번째 테이블만 선택

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

# XLSX 파일 생성 및 저장
xlsxFile = htmlFile.replace('.html', '.xlsx')
df.to_excel(xlsxFile, sheet_name='매출장', index=False)

# HTML 파일 삭제
os.remove(new_filename)

# 완료 알림 표시
print("\n" + xlsxFile + "이 생성되었습니다.")

# 프로그램 종료 시간 기록
end_time = time.time()

# 전체 실행 시간 계산 및 출력
elapsed_time = end_time - start_time
print(f"프로그램 실행 시간: {elapsed_time:.2f}초")

os.system("pause")
