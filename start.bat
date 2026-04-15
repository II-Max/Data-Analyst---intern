@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ==================================================
echo        KHỞI ĐỘNG ỨNG DỤNG - PHIÊN BẢN HOÀN CHỈNH
echo ==================================================
echo.

:: 1. Kiem tra xem Python da cai chua
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [X] LỖI: Không tìm thấy Python. Vui lòng check "Add to PATH" khi cài.
    pause
    exit /b
)
echo [V] Python đã sẵn sàng.

:: 2. Chi tao venv neu chua ton tai (Bo dau ngoac don de tranh loi CMD)
if not exist "venv\" (
    echo [*] Đang tạo môi trường ảo venv...
    python -m venv venv
) else (
    echo [V] Môi trường ảo đã tồn tại.
)

:: 3. Kich hoat moi truong ao
call venv\Scripts\activate.bat
echo [V] Đã kích hoạt venv.

:: 4. Nang cap PIP va cai thu vien
echo [*] Đang kiểm tra và cập nhật thư viện...
python -m pip install --upgrade pip >nul 2>&1
pip install flask flask-cors pandas openpyxl python-dotenv Werkzeug --default-timeout=1000
echo [V] Hoàn tất thư viện.

:: 5. Chay ung dung
echo.
echo ==================================================
echo [*] ĐANG CHẠY ỨNG DỤNG (Nhấn Ctrl+C để tắt)
echo ==================================================
python app.py

pause