@echo off
chcp 65001 >nul
title YEV2MF - XML'den Excel'e Dönüşüm Aracı
echo YEV2MF - XML'den Excel'e Dönüşüm Aracı başlatılıyor...
echo.

:: Python'un kurulu olup olmadığını kontrol et
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo HATA: Python kurulu değil veya PATH'de bulunamadı!
    echo Lütfen Python'u https://www.python.org/ adresinden indirip kurun.
    pause
    exit /b 1
)

:: Gerekli kütüphanelerin yüklü olup olmadığını kontrol et
python -c "import tkinter, openpyxl" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Gerekli Python kütüphaneleri yükleniyor...
    python -m pip install --upgrade pip
    python -m pip install openpyxl
    if %ERRORLEVEL% NEQ 0 (
        echo HATA: Gerekli kütüphaneler yüklenirken bir hata oluştu!
        pause
        exit /b 1
    )
)

:: Uygulamayı başlat
python "%~dp0yev2mf_gui.py"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo HATA: Uygulama başlatılırken bir hata oluştu!
    pause
)