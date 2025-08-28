@echo off
echo Going API...
call conda activate DICK
cd /d "%~dp0Code"
start "API" cmd /k "call conda activate DICK && cd /d "%~dp0Code" && python api.py"

echo Going Vue...
start "Vue" cmd /k "cd /d "%~dp0vue-vben-admin" && pnpm dev:antd"

exit