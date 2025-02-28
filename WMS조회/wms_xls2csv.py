import os
import csv
from lxml import html
from tqdm import tqdm

print("인터넷주문시스템(WMS)에서 다운받은 xls파일을 csv로 변환해주는 프로그램\n같은 폴더에 파일이 있어야 작동 됨")

# html 파일 경로
filename = input("파일 이름 입력(확장자 제외)") + ".xls"
new_filename = filename.rsplit('.', 1)[0] + '.html'
os.rename(filename, new_filename)
htmlFile = new_filename

# html 파일 열기
with open(htmlFile, 'rb') as f:
    html_content = f.read()

# html 파일 구문 분석
parser = html.HTMLParser(encoding='utf-8')
root = html.fromstring(html_content, parser=parser)

# csv 파일 생성
csvFile = htmlFile.replace('.html', '.csv')

# html 파일의 테이블 객체
table = root.xpath('//table')[0]  # 첫 번째 테이블만 선택

# CSV 파일 쓰기
with open(csvFile, 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    
    # 테이블의 행과 열을 순회하며 셀에 값 쓰기
    for tr in tqdm(table.xpath('.//tr')):
        row = []
        for td in tr.xpath('.//td'):
            # 셀의 텍스트 값에서 콤마 제거
            value = td.text_content().strip().replace(',', '')
            row.append(value)
        writer.writerow(row)

# html 파일 삭제
os.remove(new_filename)

# 완료 알림 표시
print("\n" + csvFile + "이 생성되었습니다.")

os.system("pause")
