@echo off
setlocal enabledelayedexpansion

:: 變數設置
set PYTHON_ENV=C:\Users\a0987\env
set SCRIPTS=CyberTracker:UI.py web_capture:web_capture.py merge_csv:merge_csv.py csv_to_xlsx:csv_to_xlsx.py

:: 確保 Python 環境可用
if exist %PYTHON_ENV%\Scripts\activate (
    echo starting venv
    call %PYTHON_ENV%\Scripts\activate
) else (
    echo didn't find vevn, using global settting...
)

:: 開始打包
for %%A in (%SCRIPTS%) do (
    for /f "tokens=1,2 delims=:" %%B in ("%%A") do (
        echo packing %%C to %%B.exe...
        pyinstaller --onefile --noconsole %%B %%C
        if %errorlevel%==0 (
            echo %%B success
        ) else (
            echo %%B fail
        )
    )
)

echo all done!
pause
