import os

import xlsxwriter
from lxml import html
from tqdm import tqdm

print("인터넷주문시스템(WMS)에서 다운받은 xls파일을 xlsx로 변환해주는 프로그램\n")

# html 파일 경로
filename = input("File name? ")+".xls"
new_filename = filename.rsplit('.',1)[0] + '.html'
os.rename(filename, new_filename)
htmlFile = new_filename

# html 파일 열기
with open(htmlFile, 'r', encoding='utf-8') as f:
    html_content = f.read()

# html 파일 구문 분석
root = html.fromstring(html_content)

# xlsx 파일 생성
xlsxFile = htmlFile.replace('.html','.xlsx')
workbook = xlsxwriter.Workbook(xlsxFile)

# xlsx 파일의 시트 객체
worksheet = workbook.add_worksheet()

# html 파일의 테이블 객체
table = root.xpath('//table')[0] # 첫 번째 테이블만 선택

# 테이블의 행과 열을 순회하며 셀에 값 쓰기
row = 0
for tr in tqdm(table.xpath('.//tr')):
    col = 0
    for td in tr.xpath('.//td'):
        # 셀의 텍스트 값
        value = td.text_content().strip()
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
print("\n"+xlsxFile+"이 생성되었습니다.")

os.system("pause")
