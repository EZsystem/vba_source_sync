@echo off
setlocal enabledelayedexpansion

:: ======================================================================
::  Excelテーブル AIサンプルデータ生成バッチ（機密保持・一括処理版）
:: ======================================================================

set PY_EXE=D:\My_code\11_workspaces\_shared_envs\python64\Scripts\python.exe
set GEN_SCRIPT=D:\My_code\11_workspaces\VBA_manager\vba_source_sync\bin\gen_ai_samples.py
set TARGET_DIR=D:\My_code\11_workspaces\VBA_manager\vba_source_sync\workspace

echo ----------------------------------------------------------------------
echo  AIサンプルデータ生成を開始します...
echo  (生成されたデータは table_values フォルダ内に保存されます)
echo ----------------------------------------------------------------------

if not exist "%PY_EXE%" (
    echo [ERROR] Python環境が見つかりません: %PY_EXE%
    pause
    exit /b
)

:: workspaceフォルダ配下の .md ファイルを探す
for /r "%TARGET_DIR%" %%f in (*.md) do (
    set FILE_NAME=%%~nxf
    :: すでに生成済みの _sample.md は除外する
    echo !FILE_NAME! | findstr /i "_sample.md" >nul
    if errorlevel 1 (
        echo [処理中] %%~nf ...
        "%PY_EXE%" "%GEN_SCRIPT%" "%%f"
    )
)

echo.
echo ----------------------------------------------------------------------
echo  すべての処理が完了しました。
echo  ※各 table_values フォルダ内の _sample.md をご確認ください。
echo ----------------------------------------------------------------------
pause
