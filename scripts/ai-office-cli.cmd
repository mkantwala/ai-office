@echo off
setlocal
set "officeRoot=%~dp0.."
if exist "%officeRoot%\.build\node-24.18\node-v24.18.0-win-x64\node.exe" (
  "%officeRoot%\.build\node-24.18\node-v24.18.0-win-x64\node.exe" "%officeRoot%\extensions\ai-office-pdf\out\cli.cjs" %*
) else (
  node "%officeRoot%\extensions\ai-office-pdf\out\cli.cjs" %*
)
exit /b %errorlevel%
