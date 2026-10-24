@echo off
REM Tạo repo Git cho thư mục này và commit lần đầu. Chạy 1 lần: bấm đúp file này, hoặc gõ  tao_git.bat  trong Terminal.
chcp 65001 >nul
git init -b main
git add -A
git commit -m "Du lieu van ban phap luat: kiem ke, OCR, tach cau truc, chunking, doi chieu"
echo.
echo Xong. Day len repo cua nhom bang 2 lenh:
echo    git remote add origin ^<link repo cua nhom^>
echo    git push -u origin main
pause
