@echo off
title Launcher Aplikasi Imbalan Kerja PSAK 219 - KKA VAB

echo ====================================================================
echo   MEMULAI SISTEM AKTUARIA IMBALAN PASCA KERJA KARYAWAN (PSAK 219)
echo ====================================================================
echo.

:: 🛠️ SESUAIKAN JALUR FOLDER DI BAWAH INI DENGAN LOKASI DI LAPTOP DEMO
cd /d "D:\Aryax punya\2026\KKA VAB\Aplikasi\imbalan-pasca-kerja-app"

echo [1/2] Memeriksa dan menginstal library pendukung...
pip install -r requirements.txt --quiet

echo.
echo [2/2] Menjalankan dashboard utama Streamlit...
echo Aplikasi akan otomatis terbuka di browser Anda beberapa saat lagi.
echo (Mohon jangan menutup jendela hitam ini selama aplikasi digunakan)
echo.

:: Menjalankan Streamlit secara otomatis
py -m streamlit run app.py

pause