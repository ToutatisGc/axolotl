@echo off

rem 检查是否已安装docsify-cli
where docsify >nul 2>nul
if %errorlevel% neq 0 (
  echo Error: [docsify-cli]未安装,请使用[npm install docsify-cli -g]安装.
  exit /b 1
)

docsify serve ../..