@echo off
chcp 65001 >nul
echo ====================================
echo  Excel 工具集 - 启动Web版
echo ====================================
echo.

pip install streamlit pandas openpyxl -q

echo 启动中...
echo 启动后手机和电脑在同一WiFi下都可访问
echo.
streamlit run app.py --server.port 8501
pause