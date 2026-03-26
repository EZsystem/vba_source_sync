@echo off
setlocal
chcp 65001 > nul

:: --- ここを仮想環境のパスに変更 ---
set PYTHON_EXE="D:\My_code\11_workspaces\VBA_manager\.venv\Scripts\python.exe"

set SCRIPT_PATH="D:\My_code\11_workspaces\VBA_manager\vba_source_sync\src\extract_sql.py"
set TARGET_DB="D:\My_DataBase\RN_ProjectSummary.accdb"

echo ============================================
echo  Access SQL資産 抽出プロセスを開始します (venv使用)
echo ============================================

:: 仮想環境のPythonで実行
%PYTHON_EXE% %SCRIPT_PATH% %TARGET_DB%

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] 処理に失敗しました。
) else (
    echo.
    echo [SUCCESS] 正常に終了しました。
)

pause