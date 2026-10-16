@echo off
chcp 936 >nul
setlocal
set "OUT=%~dp0诊断结果.txt"
set "TARGET=%~1"

if "%TARGET%"=="" (
  echo.
  echo 请把报错里那个 xlsx 文件用鼠标拖到本窗口，然后按回车：
  echo.
  set /p TARGET=
)
set "TARGET=%TARGET:"=%"

if not exist "%TARGET%" (
  echo.
  echo 找不到这个文件：%TARGET%
  pause
  exit /b 1
)

for %%I in ("%TARGET%") do set "FOLDER=%%~dpI"
set "F2=%FOLDER:~0,-1%"
for %%I in ("%F2%") do set "PARENT=%%~dpI"

> "%OUT%" echo ===== 查寝计算器「写入副本失败」诊断报告 =====
>>"%OUT%" echo 时间: %date% %time%
>>"%OUT%" echo 失败的文件: %TARGET%
>>"%OUT%" echo 所在文件夹: %FOLDER%

>>"%OUT%" echo.
>>"%OUT%" echo --- 0. 这个盘是什么盘（FileSystemType 应是 NTFS；exFAT/FAT32/CDFS 会有各种限制）---
>>"%OUT%" powershell -NoProfile -ErrorAction Stop -Command "try{Get-Volume | Format-Table DriveLetter,FileSystemType,DriveType,HealthStatus -AutoSize | Out-String -Width 200}catch{'查询不到'}" 2>&1

>>"%OUT%" echo.
>>"%OUT%" echo --- 1. 当前登录的 Windows 账户 ---
>>"%OUT%" whoami 2>&1

>>"%OUT%" echo.
>>"%OUT%" echo --- 2. 文件属性（行首出现 R 就是「只读」）---
>>"%OUT%" echo [失败的那个副本文件]:
>>"%OUT%" attrib "%TARGET%" 2>&1
>>"%OUT%" echo [它所在文件夹里的所有文件]:
>>"%OUT%" attrib "%FOLDER%*.*" 2>&1
>>"%OUT%" echo [上一级文件夹里的原件]:
>>"%OUT%" attrib "%PARENT%*.xlsx" 2>&1

>>"%OUT%" echo.
>>"%OUT%" echo --- 3. 失败文件自己的权限 ---
>>"%OUT%" icacls "%TARGET%" 2>&1

>>"%OUT%" echo.
>>"%OUT%" echo --- 4. 所在「副本」文件夹的权限 ---
>>"%OUT%" icacls "%F2%" 2>&1

>>"%OUT%" echo.
>>"%OUT%" echo --- 5. 上一级文件夹的权限 ---
>>"%OUT%" icacls "%PARENT%." 2>&1

>>"%OUT%" echo.
>>"%OUT%" echo --- 6. 实测：在这个文件夹里能不能新建文件 ---
echo test > "%FOLDER%_写测试.txt" 2>nul
if not exist "%FOLDER%_写测试.txt" goto :sec6fail
>>"%OUT%" echo 新建文件: 成功（说明文件夹本身可以写）
goto :sec7
:sec6fail
>>"%OUT%" echo 新建文件: 失败（连新建都不行 = 文件夹权限问题）
:sec7

>>"%OUT%" echo.
>>"%OUT%" echo --- 7. 实测：把文件设成只读后还能不能往里写（验证「只读」是否会造成这个报错）---
echo a > "%FOLDER%_只读测试.txt" 2>nul
attrib +R "%FOLDER%_只读测试.txt" >nul 2>&1
echo b >> "%FOLDER%_只读测试.txt" 2>nul
>>"%OUT%" echo 写入后这个测试文件的内容：
>>"%OUT%" type "%FOLDER%_只读测试.txt" 2>&1
>>"%OUT%" echo 判断：只有 a = 只读确实会挡住写入；a 和 b 都有 = 只读挡不住，原因在别处

>>"%OUT%" echo.
>>"%OUT%" echo --- 8. 实测：重新复制一份（cmd 的 copy 不继承只读属性），看它能不能写 ---
cmd /c copy /y "%TARGET%" "%FOLDER%_复制测试.xlsx" >nul 2>&1
if not exist "%FOLDER%_复制测试.xlsx" goto :sec8fail
for %%A in ("%FOLDER%_复制测试.xlsx") do set "S1=%%~zA"
>>"%OUT%" echo 新复制件的属性:
>>"%OUT%" attrib "%FOLDER%_复制测试.xlsx" 2>&1
echo x >> "%FOLDER%_复制测试.xlsx" 2>nul
for %%A in ("%FOLDER%_复制测试.xlsx") do set "S2=%%~zA"
>>"%OUT%" echo 写入前 %S1% 字节，写入后 %S2% 字节（变大了 = 能写；没变 = 被拒绝）
goto :sec9
:sec8fail
>>"%OUT%" echo 复制失败
:sec9

>>"%OUT%" echo.
>>"%OUT%" echo --- 9. Windows「受控文件夹访问」是否开启（1=True 开启，0=False 关闭）---
>>"%OUT%" powershell -NoProfile -ErrorAction Stop -Command "try{(Get-MpPreference).EnableControlledFolderAccess}catch{'查询不到'}" 2>&1

>>"%OUT%" echo.
>>"%OUT%" echo ===== 报告结束 =====

attrib -R "%FOLDER%_只读测试.txt" >nul 2>&1
del /f /q "%FOLDER%_写测试.txt" >nul 2>&1
del /f /q "%FOLDER%_只读测试.txt" >nul 2>&1
del /f /q "%FOLDER%_复制测试.xlsx" >nul 2>&1

echo.
echo 诊断完成，结果已保存到：
echo %OUT%
echo 请把这个文件发给 AI（或者把内容整段贴过来）。
echo.
pause
