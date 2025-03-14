import os
from PyPDF2 import PdfReader, PdfWriter

def reduce_margins(input_pdf, output_pdf, top=0, bottom=0, left=0, right=0):
    with open(input_pdf, 'rb') as in_file:
        reader = PdfReader(in_file)
        writer = PdfWriter()

        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            # Adjust crop box margins
            page.cropbox.lower_left = (left, bottom)
            page.cropbox.upper_right = (
                page.mediabox.width - right,
                page.mediabox.height - top
            )
            writer.add_page(page)

        with open(output_pdf, 'wb') as out_file:
            writer.write(out_file)

def process_all_pdfs_in_folder(folder_path, top=0, bottom=0, left=0, right=0):
    # List all PDF files in the folder
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    
    for pdf_file in pdf_files:
        input_path = os.path.join(folder_path, pdf_file)
        output_path = os.path.join(folder_path, f"{os.path.splitext(pdf_file)[0]}_result.pdf")
        
        print(f"Processing: {pdf_file} -> {os.path.basename(output_path)}")
        reduce_margins(input_path, output_path, top, bottom, left, right)

if __name__ == "__main__":
    # Get user inputs for margins
    print ("권장 : 35")
    left_margin = int(input("왼쪽 여백 (포인트): "))
    right_margin = int(input("오른쪽 여백 (포인트): "))
    top_margin = 0  # 상단 여백 (필요시 수정 가능)
    bottom_margin = 0  # 하단 여백 (필요시 수정 가능)

    # Process all PDFs in the current folder
    current_folder = os.getcwd()
    process_all_pdfs_in_folder(current_folder, top_margin, bottom_margin, left_margin, right_margin)
