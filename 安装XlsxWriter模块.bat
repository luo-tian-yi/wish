@echo off
chcp 65001
echo.
echo python版本
python -V
echo.
echo ------------------------------
echo 正在安装XlsxWriter模块
pip install XlsxWriter
echo.
echo 安装完成
echo.
echo ------------------------------
echo 查看模块
pip list
echo ------------------------------
pause