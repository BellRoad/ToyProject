@echo off
if "%~1"=="" (
    echo ����: pyinst "�����̸�"
    exit /b 1
)
pyinstaller -F -w --upx-dir g:\Study\python\upx %1
