#카카오톡 친구 목록 삭제를 위해 작성

import pyautogui
import time

# 숫자 입력 받기
num = int(input("반복 횟수를 입력하세요: "))

# 마우스 좌클릭 후 엔터 입력 함수
def click_and_enter():
    pyautogui.click()  # 마우스 좌클릭
    time.sleep(0.5)
    pyautogui.press('down', presses=2)
    pyautogui.press('enter')  # 엔터 입력

# 입력받은 횟수만큼 반복
for _ in range(num):
    click_and_enter()
    time.sleep(0.5)