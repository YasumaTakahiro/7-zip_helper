@echo off
chcp 65001 >nul
:: PowerShellスクリプトを実行（UTF-8で読み込み）
powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "$scriptPath = '%~dp0unzip_to_extract.ps1'; $content = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8); [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; $OutputEncoding = [System.Text.Encoding]::UTF8; $scriptBlock = [scriptblock]::Create($content); & $scriptBlock '%~1'"
pause
