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

while True:
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

# 엣지 드라이버 실행행
service = Service(executable_path=EdgeChromiumDriverManager().install())
driver = EdgeDriver(service=service)

# 엣지 드라이버에 url 주소 넣고 실행
driver.get('http://order.mydongsim.com/Login.do?user_id=finance3&password=6093')

# 페이지가 완전히 로딩되도록 2초동안 기다림
time.sleep(2)

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

# 법인, 센터 구분하여 파일을 다운로드
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

    down_url = "http://order.mydongsim.com/Report/Report090excel.jsp?comp_cd=" + compcd + "&date_from=" + datefrom + "&date_to=" + dateto + "&agy_name=" + agy_name + "&agency=" + agency + "&goods_name=&goods=&rule_cd1=&rule_cd2=&rule_cd3=&badgb=1"
    driver.get(down_url)

    # 재고리스트(누계) 파일 다운로드 되는 동안 프로그램 종료 대기
    filename = "재고리스트(누계).xls"
    file_path = os.path.join(downloads_folder, filename)
    while not os.path.exists(file_path):
        time.sleep(3)

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
    new_filename = os.path.join(downloads_folder, filename.rsplit('.', 1)[0] + ' ' + compcd + ' ' + agy_name + ' ' + datefrom + '-' + dateto + '.html')
    os.rename(file_path, new_filename)

# 파일 다운로드 완료 후 엣지드라이버 종료
driver.quit()


# 작업 폴더 변경
os.chdir(downloads_folder)

# 파일을 pandas로 변환 후 재고 센터를 병합
import pandas as pd
import re

def process_files(file_pattern):
    files = [f for f in os.listdir('.') if re.match(file_pattern, f)]
    
    if not files:
        return None, None, None

    dfs = []
    for file in files:
        df = pd.read_html(file)[0]
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)
        dfs.append(df)

    merged_df = pd.concat(dfs, ignore_index=True)
    merged_df = merged_df.drop('순번', axis=1)

    numeric_columns = merged_df.columns.drop(['품목코드', '품목명'])
    merged_df[numeric_columns] = merged_df[numeric_columns].apply(pd.to_numeric, errors='coerce')

    merged_df = merged_df.sort_values(['품목코드', '품목명'])
    grouped_df = merged_df.groupby(['품목코드', '품목명']).sum().reset_index()

    grouped_df['출고 합계'] = grouped_df['출고 수량'] + grouped_df['반품 수량'] + grouped_df['A/S 수량'] + grouped_df['조정 수량'] + grouped_df['가공 수량']

    new_order = ['품목코드', '품목명', '전재고 수량', '전재고 금액', '입고 수량', '출고 합계', '현재고 수량', '현재고 금액', '출고 수량', '반품 수량', 'A/S 수량', '조정 수량', '가공 수량', '입고 금액', '출고 금액', '반품 금액', 'A/S 금액', '조정 금액', '가공 금액']
    grouped_df = grouped_df[new_order]

    date_match = re.search(r'(\d{8}-\d{8})', files[0])
    date_range = date_match.group(1) if date_match else ''

    return grouped_df, date_range, files[0]

patterns = [
    r'재고리스트\(누계\) 5월5일 (교원|안성) \d{8}-\d{8}\.html',
    r'재고리스트\(누계\) 동심 (교원|안성) \d{8}-\d{8}\.html',
    r'재고리스트\(누계\) 킨더 (교원|안성) \d{8}-\d{8}\.html'
]

for pattern in patterns:
    df, date_range, sample_file = process_files(pattern)
    if df is not None:
        output_name = sample_file.replace('.html', '.xlsx')
        output_name = re.sub(r' (교원|안성)', ' 통합', output_name)
        
        df.to_excel(output_name, index=False)
        print(f"{output_name} 파일이 생성되었습니다.")
    else:
        print(f"{pattern}에 해당하는 파일을 찾을 수 없습니다.")

# 작업이 완료된 후 HTML 파일 삭제
def delete_html_files(patterns):
    for pattern in patterns:
        files = [f for f in os.listdir('.') if re.match(pattern, f)]
        for file in files:
            try:
                os.remove(file)
                print(f"{file} 삭제 완료")
            except Exception as e:
                print(f"{file} 삭제 실패: {e}")

delete_html_files(patterns)

# 프로그램 종료 시간 기록
end_time = time.time()

# 전체 실행 시간 계산 및 출력
elapsed_time = end_time - start_time
print(f"프로그램 실행 시간: {elapsed_time:.2f}초")

os.system("pause")