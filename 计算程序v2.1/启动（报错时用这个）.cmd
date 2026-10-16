@echo off
chcp 65001 >nul
rem 备用启动器：如果直接双击 exe 出现 “Could not create temporary directory!”
rem 就改双击本文件，它会先指定一个可用的临时目录再启动程序。

set "TMPDIR=%~dp0_calc_tmp"
if not exist "%TMPDIR%" mkdir "%TMPDIR%"
if not exist "%TMPDIR%" set "TMPDIR=%LOCALAPPDATA%\Temp\查寝总分计算器"
if not exist "%TMPDIR%" mkdir "%TMPDIR%" >nul 2>nul

set "TEMP=%TMPDIR%"
set "TMP=%TMPDIR%"

start "" "%~dp0查寝总分计算器.exe" %*
