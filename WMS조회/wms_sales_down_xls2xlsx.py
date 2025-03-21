# selenium의 webdriver를 사용하기 위한 import
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service

# selenium으로 무엇인가 입력하기 위한 import
from selenium.webdriver.common.by import By

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

# 입력 날짜 형식 변경
def datetrans(date):
    year = date[:4]
    month = date[4:6]
    day = date[6:]
    return f"{year}.{month}.{day}"

datetrans(datefrom)
datetrans(dateto)

# 웹드라이버 버젼 관련 WARNING 레벨 이상의 메시지 숨기기
import logging

class MyHandler(logging.Handler):
    def emit(self, record):
        if "The msedgedriver version" not in record.msg:
            logging._log(record.levelno, record.msg, record.args, record.exc_info, record.filename, record.lineno, record.funcName, record.created, record.msecs, record.relativeCreated, record.thread, record.threadName, record.process, record.processName, record.stack_info)

logging.basicConfig(level=logging.WARNING, handlers=[MyHandler()])

# SSL 관련 오류 안보이게 하고 엣지 드라이버 실행
options = webdriver.EdgeOptions()
options.add_argument("disable-gpu")
options.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36")
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver_path = "G:\Study\ToyProject\msedgedriver.exe"
service = Service(executable_path=driver_path)
driver = webdriver.Edge(service=service, options=options)

# 엣지 드라이버에 url 주소 넣고 실행
driver.get('http://order.mydongsim.com/Login.do?user_id=finance3&password=6093')

# 페이지가 완전히 로딩되도록 2초동안 기다림
time.sleep(2)

'''
# ID, PW 창을 찾아서 입력
driver.find_element(By.XPATH, '//*[@id="user_id"]').send_keys('finance3')
driver.find_element(By.XPATH, '//*[@id="password"]').send_keys('6093')
driver.find_element(By.XPATH, '//*[@id="loginBtn"]/img').click()
time.sleep(2)
''''

down_url = "http://order.mydongsim.com/Report/Report070excel.jsp?comp_cd=" + compcd + "&date_from=" + datefrom + "&date_to=" + dateto + "&supply_cust_nm=&supply_custno=&goods_name=&goods=&rule_cd6="
driver.get(down_url)

# 매출장 파일 다운로드 되는 동안 프로그램 종료 대기
filename = "매출장.xls"
while not os.path.exists(os.path.join(filename)):
    time.sleep(5)

# 파일 다운로드 완료 후 엣지드라이버 종료
driver.quit()

print("xls를 xlsx로 변환 시작")

# 액셀 관련 모듈 import
import xlsxwriter
from lxml import html
from tqdm import tqdm

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

# xls를 html로 이름 변환
new_filename = filename.rsplit('.', 1)[0] + ' ' + compcd + ' ' + datefrom + '-' + dateto + '.html'
os.rename(filename, new_filename)
htmlFile = new_filename

# html 파일 열기
with open(htmlFile, 'r', encoding='utf-8') as f:
    html_content = f.read()

# html 파일 구문 분석
root = html.fromstring(html_content)

# xlsx 파일 생성
xlsxFile = htmlFile.replace('.html', '.xlsx')
workbook = xlsxwriter.Workbook(xlsxFile)

# xlsx 파일의 시트 객체
worksheet = workbook.add_worksheet('매출장')

# html 파일의 테이블 객체
table = root.xpath('//table')[0]  # 첫 번째 테이블만 선택

# 테이블의 행과 열을 순회하며 셀에 값 쓰기
row = 0
for tr in tqdm(table.xpath('.//tr')):
    col = 0
    for td in tr.xpath('.//td'):
        # 셀의 텍스트 값
        value = td.text_content().strip()

        # C, D 열의 점(.)을 하이픈(-)으로 변경
        if col == 2 or col == 3:
            value = value.replace('.', '-')

        # A, I, M, N, O 열의 문자를 숫자로 변경
        if col == 0 or col == 8 or col == 12 or col == 13 or col == 14:
            try:
                value = int(value.replace(',', ''))
            except ValueError:
                pass

        # 셀의 배경색 값
        bgcolor = td.get('bgcolor')
        # 셀의 서식 객체
        format = workbook.add_format()
        # 셀의 배경색 설정
        if bgcolor:
            format.set_bg_color(bgcolor)
        # xlsx 파일의 셀에 값과 서식 쓰기
        worksheet.write(row, col, value, format)
        col += 1
    row += 1

# xlsx 파일 닫기
workbook.close()

# html 파일 삭제
os.remove(new_filename)

# 완료 알림 표시
print("\n" + xlsxFile + "이 생성되었습니다.")

os.system("pause")
