@echo off
chcp 65001 >nul
cd /d "D:\workplace\drug_trace_app"
echo 药品追溯码 v3.0 一键发布
echo 启动中...
powershell -ExecutionPolicy Bypass -File "publish.ps1"
pause