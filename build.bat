@echo off
chcp 65001 >nul
go build -o ../AddToPath.exe
if errorlevel 1 (
    echo 编译失败！
    pause
    exit /b 1
) else (
    echo 编译成功！
)
