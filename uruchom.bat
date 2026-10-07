@echo off
chcp 65001 >nul
cd /d "%~dp0"
python obliczenia2word.py %*
pause
