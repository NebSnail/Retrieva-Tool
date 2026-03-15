@echo off
setlocal
cd /d "%~dp0"

set "PY=%~dp0.venv\Scripts\python.exe"

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "通用件查询工具.spec" del /q "通用件查询工具.spec"

"%PY%" -m PyInstaller --noconfirm --clean --windowed --name "通用件查询工具" --icon "tool.png" --add-data "tool.png;." --add-data "types_config.json;." "query_tool.py"

if %errorlevel% neq 0 (
  echo 打包失败，请检查上方日志。
  exit /b %errorlevel%
)

echo 打包完成: dist\通用件查询工具\通用件查询工具.exe
endlocal
