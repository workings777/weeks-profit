@echo off
cd /d "C:\Users\FURSYS\Desktop\kyun\weeks_profit"

:: server.py 백그라운드 실행
start "" /min "C:\Users\FURSYS\AppData\Local\Python\bin\python.exe" server.py

:: 서버 기동 대기
timeout /t 2 /nobreak >nul

:: 브라우저 열기
start "" "http://localhost:8080"
