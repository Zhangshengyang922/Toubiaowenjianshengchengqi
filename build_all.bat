:: ═══════════════════════════════════════════════════════
::  一键构建脚本 —— 打包 + 构建安装包
::
::  使用前请确保:
::    1. pip install pyinstaller
::    2. pip install -r requirements-gui.txt
::    3. 已安装 Inno Setup (https://jrsoftware.org/isdl.php)
::    4. 修改 version.py 中的版本号和仓库地址
::    5. (可选) 在项目根目录放置 icon.ico 作为应用图标
::
::  使用方法: 双击运行此文件，或在命令行运行
:: ═══════════════════════════════════════════════════════

@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo ═══════════════════════════════════════════════════════
echo   招投标文件自动生成系统 —— 构建打包
echo ═══════════════════════════════════════════════════════
echo.

:: ── 步骤1: 检查依赖 ──
echo [1/4] 检查 Python 依赖...
pip install -r requirements-gui.txt --quiet
if %errorlevel% neq 0 (
    echo   依赖安装失败，请检查网络连接
    pause
    exit /b 1
)
echo   依赖检查完成

:: ── 步骤2: 读取版本号 ──
for /f "tokens=2 delims= " %%v in ('python -c "from version import __version__; print(__version__)"') do set VERSION=%%v
echo [2/4] 当前版本: v%VERSION%

:: ── 步骤3: PyInstaller 打包 ──
echo [3/4] PyInstaller 打包 (需要几分钟)...
python build.py
if %errorlevel% neq 0 (
    echo   PyInstaller 打包失败！
    pause
    exit /b 1
)

:: ── 步骤4: Inno Setup 构建安装包 ──
echo [4/4] 构建安装包...
echo.

:: 自动更新 installer.iss 中的版本号
python -c "
import re
iss_path = 'installer.iss'
with open(iss_path, 'r', encoding='utf-8') as f:
    content = f.read()

from version import __version__
exe_name = f'BiddingDocGen_v{__version__}.exe'
content = re.sub(r'#define MyAppVersion \".*\"', f'#define MyAppVersion \"{__version__}\"', content)
content = re.sub(r'#define MyAppExeName \".*\"', f'#define MyAppExeName \"{exe_name}\"', content)

with open(iss_path, 'w', encoding='utf-8') as f:
    f.write(content)
print(f'  installer.iss 已更新为 v{__version__}')
"

:: 查找 Inno Setup 编译器
set ISCC=
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
)
if exist "C:\Program Files\Inno Setup 6\ISCC.exe" (
    set "ISCC=C:\Program Files\Inno Setup 6\ISCC.exe"
)

if "%ISCC%"=="" (
    echo.
    echo ═══════════════════════════════════════════════════════
    echo   Inno Setup 未找到！
    echo.
    echo   请从以下地址下载安装 Inno Setup:
    echo   https://jrsoftware.org/isdl.php
    echo.
    echo   安装后可再次运行此脚本，或手动用 Inno Setup
    echo   Compiler 打开 installer.iss 编译。
    echo ═══════════════════════════════════════════════════════
    echo.
    echo ✓ PyInstaller .exe 已生成在 dist\ 目录
    explorer dist
    pause
    exit /b 0
)

"%ISCC%" installer.iss
if %errorlevel% neq 0 (
    echo   安装包构建失败！
    pause
    exit /b 1
)

echo.
echo ═══════════════════════════════════════════════════════
echo   ✓ 构建完成！
echo.
echo   输出文件:
echo     - 程序: dist\BiddingDocGen_v%VERSION%.exe
echo     - 安装包: dist\BiddingDocGen_Setup_v%VERSION%.exe
echo ═══════════════════════════════════════════════════════
echo.

explorer dist
pause
