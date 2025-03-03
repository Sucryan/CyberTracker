@echo off
setlocal enabledelayedexpansion

:: 變數設置
set PYTHON_ENV=C:\pyenv\myenv
set SCRIPTS=CyberTracker:UI.py web_capture:web_capture.py merge_csv:merge_csv.py csv_to_xlsx:csv_to_xlsx.py

:: 確保 Python 環境可用
if exist %PYTHON_ENV%\Scripts\activate (
    echo 啟動虛擬環境...
    call %PYTHON_ENV%\Scripts\activate
) else (
    echo 未偵測到虛擬環境，使用全局 Python...
)

:: 開始打包
for %%A in (%SCRIPTS%) do (
    for /f "tokens=1,2 delims=:" %%B in ("%%A") do (
        echo 正在打包 %%C 為 %%B.exe...
        pyinstaller --onefile --name %%B %%C
        if %errorlevel%==0 (
            echo %%B 打包成功!
        ) else (
            echo %%B 打包失敗!
        )
    )
)

echo 所有腳本打包完成!
pause
