@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   배송 송장 주소 자동 정제 프로그램
echo ============================================
where python >nul 2>nul
if errorlevel 1 (
  echo Python 이 설치되어 있지 않습니다. https://www.python.org 에서 설치 후 다시 실행하세요.
  pause
  exit /b
)
python -c "import openpyxl" 2>nul || python -m pip install openpyxl
python "%~dp0주소정제.py" %*
echo.
pause
