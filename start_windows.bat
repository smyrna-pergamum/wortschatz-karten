@echo off
cd /d %~dp0
python -m pip install -r requirements.txt
python wortschatz_karten.py
pause
