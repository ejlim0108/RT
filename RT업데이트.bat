@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not exist "%PY%" set "PY=python"
echo 최신 캉골RT내역_YYMMDD.xlsx 로 RT 검토맵을 갱신하고 GitHub에 올립니다.
echo.
"%PY%" -I build_rt.py --push
echo.
pause
