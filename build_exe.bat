@echo off
setlocal
cd /d "%~dp0"

set "PY=%~dp0.venv\Scripts\python.exe"

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "检索工具.spec" del /q "检索工具.spec"

"%PY%" -m PyInstaller --noconfirm --clean --windowed --name "检索工具" --icon "logo.png" --add-data "logo.png;." --add-data "types_config.json;." "retrieval_tool.py"

if %errorlevel% neq 0 (
  echo 打包失败，请检查上方日志。
  exit /b %errorlevel%
)

echo 打包完成: dist\检索工具\检索工具.exe
endlocal
