@echo off
chcp 65001 >nul
cd /d "%~dp0"
start "" pythonw photo_album_app.py
