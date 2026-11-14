@echo off
REM Chạy thử trên Windows: cần Python 3.12 (python.org). Lần đầu tạo môi trường và cài thư viện.
cd /d "%~dp0"
if not exist .venv (
  py -3.12 -m venv .venv
  .venv\Scripts\python -m pip install -r requirements.txt
)
.venv\Scripts\python -m ui.giao_dien
