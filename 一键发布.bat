@echo off
pushd D:\workplace\drug_trace_app
echo Starting publish...
powershell -ExecutionPolicy Bypass -NoProfile -File publish.ps1
popd
pause