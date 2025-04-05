@echo off
:: 切換到程式所在的目錄
cd /d "%~dp0"
echo Current directory is: %CD%

setlocal enabledelayedexpansion

:: 設定 Python 虛擬環境路徑
set PYTHON_ENV=E:\env

:: 定義 (exe檔名 : py檔名) 的對應清單
set SCRIPTS=CyberTracker:UI.py web_capture:web_capture.py merge_csv:merge_csv.py csv_to_xlsx:csv_to_xlsx.py

:: 檢查並啟用虛擬環境
if exist "%PYTHON_ENV%\Scripts\activate" (
    echo starting venv
    call "%PYTHON_ENV%\Scripts\activate"
) else (
    echo didn't find venv, using global setting...
)

:: 開始打包
for %%A in (%SCRIPTS%) do (
    for /f "tokens=1,2 delims=:" %%B in ("%%A") do (
        echo packing %%C to %%B.exe...
        echo Running: pyinstaller --onefile --noconsole "%%C" --name "%%B" --log-level=DEBUG
        pyinstaller --onefile --noconsole "%%C" --name "%%B" --log-level=DEBUG

        if !errorlevel!==0 (
            echo %%B success
        ) else (
            echo %%B fail
        )
    )
)

echo all done!
pause
