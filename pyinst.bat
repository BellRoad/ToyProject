@echo off
if "%~1"=="" (
    echo 사용법: pyinst "파일이름"
    exit /b 1
)
pyinstaller -F -w --upx-dir g:\Study\python\upx %1
