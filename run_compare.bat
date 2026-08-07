@echo off
chcp 65001 > nul
REM ============================================================
REM  실행 방법: 이 파일 위에 .hwp 파일 1개와 .xls(또는 .xlsx) 파일
REM  1개, 총 2개를 마우스로 함께 끌어다 놓으세요(드래그&드롭).
REM  더블클릭만 하면 파일 선택 창이 뜹니다.
REM ============================================================
cd /d "%~dp0"

REM venv가 있으면 자동으로 활성화 (더블클릭 시에도 패키지 인식되도록)
if exist "venv\Scripts\activate.bat" (
    call "venv\Scripts\activate.bat"
)

python compare_hwp_excel.py --out-dir ./output %*
pause