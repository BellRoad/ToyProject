'''
구글 블로그의 백업XML파일을 HTML로 변환
제목, 날짜, 링크, 본문의 내용을 포함
Version 1.0
'''

import xml.etree.ElementTree as ET
from html import escape
import os

def parse_blogger_xml(xml_file, output_file):
    # XML 파일 파싱
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # 네임스페이스 정의
    ns = {'atom': 'http://www.w3.org/2005/Atom'}

    # HTML 파일 시작
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('<html><head><meta charset="utf-8"><title>Blog Backup</title></head><body>')

        # 각 포스트 처리
        for entry in root.findall('atom:entry', ns):
            if entry.find('atom:category[@term="http://schemas.google.com/blogger/2008/kind#post"]', ns) is not None:
                title = entry.find('atom:title', ns).text
                date = entry.find('atom:published', ns).text
                url = entry.find('atom:link[@rel="alternate"]', ns).attrib['href']
                content = entry.find('atom:content', ns).text

                # HTML에 정보 추가
                f.write(f'<h2>{escape(title)}</h2>')
                f.write(f'<p>날짜: {escape(date)}</p>')
                f.write(f'<p>URL: <a href="{escape(url)}" target=_blank>{escape(url)}</a></p>')
                f.write(f'<div>{content}</div>')
                f.write('<hr>')

        # HTML 파일 종료
        f.write('</body></html>')

    print(f"백업이 완료되었습니다. 파일명: {output_file}")

# 사용자로부터 입력 받기
print ("구글 블로그 blogger.com의 백업XML파일을 HTML파일로 변환해줍니다.")
print ("제목, 날짜, 링크, 본문을 포함합니다.")
filename = input("XML 파일의 이름을 입력하세요(확장자 제외):")
xml_file = filename + ".xml"
output_file = filename + ".html"

# 함수 실행
parse_blogger_xml(xml_file, output_file)

os.system('pause')