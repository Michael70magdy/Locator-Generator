@echo off
cd /d "%~dp0"
py -m pip install beautifulsoup4
py locator_generator.py
pause
