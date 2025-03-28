import winreg
import os

def get_download_folder():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
            downloads_path = winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
        return downloads_path
    except Exception:
        # 레지스트리에서 찾지 못한 경우 기본 경로 반환
        return os.path.join(os.path.expanduser('~'), 'Downloads')

# 다운로드 폴더 경로 출력
print(f"다운로드 폴더 위치: {get_download_folder()}")

file_path = os.path.join(get_download_folder(), 'test.txt')
file_path_new = os.path.join(get_download_folder(), 'test2.txt')

os.rename(file_path, file_path_new)

print(file_path_new)