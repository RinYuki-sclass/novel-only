@echo off
setlocal enabledelayedexpansion

echo.
echo ==========================================
echo 🛠️  DANG CHAY RE-QC REFINE (CHINH SUA LAI)
echo ==========================================
echo.
echo * Luu y: Buoc nay se lay file 'output/vi_final.txt' hien tai
echo * lam input de AI tu chinh sua (Polish) theo cac luat moi.
echo.
echo [1/1] Dang goi main.py voi che do --mode refine...

python scripts/main.py --mode refine

echo.
echo ==========================================
echo ✅ Hoan tat! Ket qua da duoc cap nhat tai output/vi_final.txt
echo Nhan phim bat ky de thoat.
pause > nul
