@echo off
chcp 65001 >nul
:: PowerShellスクリプトを実行（UTF-8で読み込み）
:: 設定はJSON設定ファイル（unzip_to_extract_config.json）から読み込まれます。
powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "$scriptPath = '%~dp0unzip_to_extract.ps1'; $scriptDir = '%~dp0'; $content = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8); [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; $OutputEncoding = [System.Text.Encoding]::UTF8; $scriptBlock = [scriptblock]::Create($content); & $scriptBlock $scriptDir '%~1'"
:: エラー時はpauseで待機、正常終了時はPowerShellスクリプト内で5秒待機してから終了
if errorlevel 1 pause
