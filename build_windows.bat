@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements.txt
if errorlevel 1 exit /b 1
python download_models.py
if errorlevel 1 exit /b 1
python -m PyInstaller --noconfirm --clean --onefile --windowed --name PhotoArchiveCatalog --icon "assets\program_icon.ico" --add-data "models;models" --add-data "assets;assets" --hidden-import cv2 --hidden-import numpy --hidden-import openpyxl --hidden-import docx --hidden-import pyodbc --hidden-import win32com.client --hidden-import pythoncom --hidden-import pywintypes --collect-all openpyxl --collect-submodules win32com main.py
if errorlevel 1 exit /b 1
python -m PyInstaller --noconfirm --clean --onefile --windowed --uac-admin --name PhotoArchiveCatalogUpdater --icon "assets\program_icon.ico" updater.py
endlocal
