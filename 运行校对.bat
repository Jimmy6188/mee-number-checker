@echo off
chcp 65001 >nul
rem ============================================================
rem MEE 多国语数字校对工具 - 启动脚本
rem 用法: 把第 8、9 行的路径改成实际的指示稿和数据文件夹后双击
rem 结果输出到数据文件夹下的 _校对结果 子目录
rem ============================================================

set BASE=F:\vibeCoding\MEE-校对数字\10份MEE数字校对测试数据\10份MEE数字校对测试数据\CSH2026A0400\校对数字英文指示稿.pdf
set DIR=F:\vibeCoding\MEE-校对数字\10份MEE数字校对测试数据\10份MEE数字校对测试数据\CSH2026A0400

python "%~dp0mee_checker.py" --base "%BASE%" --dir "%DIR%"
pause
