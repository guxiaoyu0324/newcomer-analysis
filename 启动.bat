@echo off
chcp 65001 >nul
echo ==========================================
echo   新人数据自动化分析系统
echo ==========================================
echo.
echo 正在启动网页应用...
echo 浏览器会自动打开 http://localhost:8501
echo.
echo 按 Ctrl+C 可以停止服务
echo ==========================================
echo.
cd /d "%~dp0"
"C:\Users\ZYB\AppData\Local\Programs\Python\Python313\python.exe" -m streamlit run app.py --server.port 8501
pause
