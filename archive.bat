@echo off
chcp 65001 >nul
powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "$scriptPath = '%~dp0archive.ps1'; $scriptDir = '%~dp0'; $targetPath = '%~1'; $content = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8); [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; $OutputEncoding = [System.Text.Encoding]::UTF8; $scriptBlock = [scriptblock]::Create($content); & $scriptBlock $scriptDir $targetPath"

