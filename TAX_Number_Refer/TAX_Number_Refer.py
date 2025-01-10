from tkinter import *
from PublicDataReader import Nts

def search_status():
    user_input = entry.get()  # Entry 위젯에서 입력값 가져오기
    serviceKey = "ETCwR1A4E4jxUMosZ4RmaO+l950l84cG5IutsAQ8n9l0NQldZGIJkdKVp6ms/Nfxa0FmZR59gldjFDF6dW2+bw=="
    API = Nts(serviceKey)
    result = API.status([user_input])

    # 결과를 Label 위젯에 표시
    result_label.config(text=f"사업 상태: {result.get('b_stt')[0]}\n과세 유형: {result.get('tax_type')[0]}\n변경 날짜: {result.get('end_dt')[0]}")

def on_enter(event):
    button.invoke()  # 엔터 키를 누르면 조회 버튼 클릭

# 창 생성 및 크기 설정
window = Tk()
window.title("사업자등록번호 상태 조회")
window.geometry("300x200")  # 창 크기 설정

# 레이블, 엔트리, 버튼 등을 프레임에 추가
label = Label(window, text="사업자등록번호 입력:")
label.pack()

entry = Entry(window)
entry.pack()
entry.focus_set()  # 엔트리에 포커스 설정

button = Button(window, text="조회", command=search_status)
button.pack()

# 엔터 키 이벤트 바인딩
entry.bind("<Return>", on_enter)

result_label = Label(window, text="", justify=LEFT)
result_label.pack()

window.mainloop()