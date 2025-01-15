import tkinter as tk
from tkinter import scrolledtext
import pyautogui
import time

class KeyMacroGUI:
    def __init__(self, master):
        self.master = master
        master.title("키보드 매크로")
        master.geometry("250x250")
        master.resizable(False, False)

        self.create_widgets()

    def create_widgets(self):
        # 키 입력 안내 레이블
        self.label = tk.Label(self.master, text="키를 입력하세요\nenter, space, left, right, up, down 가능\nctrl+space,shift+d 가능")
        self.label.pack(pady=10)

        # 스크롤 가능한 텍스트 입력 영역
        self.text_area = scrolledtext.ScrolledText(self.master, width=50, height=10)
        self.text_area.pack(pady=5)

        # 실행 버튼
        self.run_button = tk.Button(self.master, text="매크로 실행", command=self.run_macro)
        self.run_button.pack(pady=5)

    def run_macro(self):
        keys = self.text_area.get("1.0", tk.END).strip().split("\n")
        if not keys:
            return

        # 5초 대기
        self.label.config(text="5초 후 매크로가 시작됩니다...")
        self.master.update()
        time.sleep(5)

        # 매크로 실행
        for key in keys:
            if key:
                pyautogui.press(key)
                time.sleep(0.1)

        self.label.config(text="매크로 실행 완료!")

def main():
    root = tk.Tk()
    app = KeyMacroGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
