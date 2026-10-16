# 查寝总分计算器 —— 打包脚本（在项目目录下用 PowerShell 运行本文件）
# 用法：右键本文件 → 使用 PowerShell 运行，或在本目录执行 powershell -File 打包.ps1
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$py   = Join-Path $root '查寝总分计算器.py'
$icon = Join-Path $root '程序图标.ico'

Write-Host '检查 PyInstaller...'
python -c "import PyInstaller" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host '未安装 PyInstaller，开始安装（需要联网）...'
    python -m pip install --user pyinstaller
}

$args = @(
    '--noconfirm', '--clean', '--onefile', '--windowed',
    '--name', '查寝总分计算器',
    '--distpath', $root,
    # 关键：把运行时解压目录固定成 exe 同目录下的 _calc_tmp，
    # 这样不再依赖系统 %TEMP%，可避免 “Could not create temporary directory!” 报错
    '--runtime-tmpdir', '.\_calc_tmp',
    '--workpath', (Join-Path $root '_build'),
    '--specpath', (Join-Path $root '_build')
)
if (Test-Path $icon) { $args += @('--icon', $icon) }
$args += $py

Write-Host '开始打包...'
python -m PyInstaller @args

# 清理打包中间文件，只留下 exe
$buildDir = Join-Path $root '_build'
if (Test-Path $buildDir) { Remove-Item $buildDir -Recurse -Force -ErrorAction SilentlyContinue }
Write-Host ''
Write-Host ('打包完成：' + (Join-Path $root '查寝总分计算器.exe'))
