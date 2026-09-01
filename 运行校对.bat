@echo off
chcp 65001 >nul
REM ============================================================
REM MEE 多国语数字校对工具 - 启动脚本
REM 用法: 将 英文全标红指示稿PDF 和 多国语文件夹 分别拖到此bat上
REM   参数1: 英文指示稿(红字标注版) 参数2: 多国语PDF文件夹 (可选参数3: 客户锚定指示原稿)
REM 例: 运行校对.bat 英文红字.pdf 语言文件夹 高光原稿.pdf
REM ============================================================

set BASE=%~1
set DIR=%~2
set ANCHOR=%~3

if "%BASE%"=="" (
  echo 用法: 将英文指示稿PDF拖到此bat; 参数2为多国语文件夹, 参数3(可选)为客户锚定指示原稿
  pause
  exit /b
)
if "%DIR%"=="" (
  echo 参数2: 请拖入多国语PDF文件夹
  pause
  exit /b
)

if "%ANCHOR%"=="" (
  python "%~dp0mee_checker.py" --base "%BASE%" --dir "%DIR%"
) else (
  python "%~dp0mee_checker.py" --base "%BASE%" --anchor "%ANCHOR%" --dir "%DIR%"
)
pause
