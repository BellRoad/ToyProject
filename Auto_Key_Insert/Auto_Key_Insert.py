import pyautogui
import time

def insert_key():
    keys = []
    print("키를 입력하세요. 종료하려면 빈 문자열을 입력하세요.")
    print("특수 키: up, down, left, right, tab, enter, space")
    print("단축키 조합: ctrl+c, alt+f4, ctrl+space, shift+space 등")
    
    while True:
        key = input("키 입력: ").lower()  # 소문자로 변환
        if key == "":
            break
        
        if '+' in key:
            # 단축키 조합 처리
            keys.append(tuple(key.split('+')))
        elif key in ['up', 'down', 'left', 'right', 'tab', 'enter', 'space']:
            # 특수 키 처리
            keys.append(key)
        else:
            # 일반 키 처리
            keys.append(key)
    
    return keys

def repeat_keys(keys, interval=0.1):
    for key in keys:
        if isinstance(key, tuple):
            # 단축키 조합 실행
            pyautogui.hotkey(*key)
        else:
            # 단일 키 실행
            pyautogui.press(key)
        time.sleep(interval)

# 키 입력 받기
keys = insert_key()

# 입력받은 키의 개수 출력
print(f"입력받은 키의 개수: {len(keys)}")

# 5초 대기 후 시작
print("5초 후 매크로가 시작됩니다.")
time.sleep(5)

# 입력받은 키 순서대로 실행
repeat_keys(keys)
