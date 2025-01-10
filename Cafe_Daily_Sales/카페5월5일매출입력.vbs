Set objShell = CreateObject("WScript.Shell")

' 명령 프롬프트를 열고 Python 스크립트 실행
objShell.Run "cmd /c python Cafe_Daily_Sales.py", 0, True

Set objShell = Nothing
