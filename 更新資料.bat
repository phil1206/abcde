@echo off
chcp 65001 >nul
echo 正在更新台股資料，請稍候...
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0update.ps1"
echo.
echo 更新完成！重新整理瀏覽器頁面即可看到最新資料。
pause
