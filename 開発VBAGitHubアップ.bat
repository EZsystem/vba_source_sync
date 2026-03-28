@echo off
title VBA_Source_Sync_System
chcp 65001 > nul

set PROJECT_ROOT=D:\My_code\11_workspaces\VBA_manager\vba_source_sync
set VENV_PYTHON=D:\My_code\11_workspaces\VBA_manager\.venv\Scripts\python.exe

echo ==================================================
echo   VBAソースコード抽出 ＆ GitHub同期システム
echo ==================================================
echo.
echo [状態] 32bit Python仮想環境を起動中...
echo [対象] Access / Excel ファイル選択待機...
echo.

rem プロジェクトルートへ移動
cd /d %PROJECT_ROOT%

rem 仮想環境のPythonを使用してメインスクリプトを実行
"%VENV_PYTHON%" bin\main.py

echo.
echo ==================================================
echo   処理が終了しました。
echo ==================================================
pause
