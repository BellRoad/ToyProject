# 문헌정보 교안 PDF파일의 좌우 여백 삭제

from PyPDF2 import PdfReader, PdfWriter

def reduce_margins(input_pdf, output_pdf, top=0, bottom=0, left=0, right=0):

    with open(input_pdf, 'rb') as in_file:
        reader = PdfReader(in_file)
        writer = PdfWriter()

        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            page.cropbox.lower_left = (left, bottom)
            page.cropbox.upper_right = (page.mediabox.width - right, page.mediabox.height - top)
            writer.add_page(page)

        with open(output_pdf, 'wb') as out_file:
            writer.write(out_file)

# 사용 예시
input_file = input("파일이름(.pdf제외): ") + ".pdf"
output_file = input_file.split(".")[0] + "_result.pdf"
left = int(input("왼쪽 여백 (포인트): "))
right = int(input("오른쪽 여백 (포인트): "))
top = 0
bottom = 0

reduce_margins(input_file, output_file, top, bottom, left, right)