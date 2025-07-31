@echo off
setlocal

REM 设置窗口标题
title Markdown Clipboard to Word Converter

echo Markdown to Word Converter (Clipboard Version)
echo ==============================================
echo.
echo 正在从剪贴板读取内容并转换为Word文档...
echo.

REM 运行转换脚本
python "%~dp0clipboard_to_word.py"

if %errorlevel% equ 0 (
    echo.
    echo 转换完成！
) else (
    echo.
    echo 转换失败，请检查错误信息。
)

echo.
echo 按任意键退出...
pause >nul
