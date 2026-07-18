[CmdletBinding()]
param(
    [switch]$Development
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$VenvDir = Join-Path $ProjectRoot ".venv"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"

if (Get-Command py -ErrorAction SilentlyContinue) {
    & py -3 -m venv $VenvDir
}
elseif (Get-Command python -ErrorAction SilentlyContinue) {
    & python -m venv $VenvDir
}
else {
    throw "没有找到 Python。请先安装 Python 3.11 或更高版本，并勾选 Add Python to PATH。"
}

if ($Development) {
    $RequirementsFile = Join-Path $ProjectRoot "requirements-windows-dev.txt"
}
else {
    $RequirementsFile = Join-Path $ProjectRoot "requirements-windows.txt"
}

& $VenvPython -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) {
    throw "pip 升级失败。"
}
& $VenvPython -m pip install -r $RequirementsFile
if ($LASTEXITCODE -ne 0) {
    throw "依赖安装失败：$RequirementsFile"
}
& $VenvPython -c "import pdfplumber, pypdfium2, PIL, docx, pyautogui, win32com.client"
if ($LASTEXITCODE -ne 0) {
    throw "Windows 运行依赖导入失败。"
}

if ($Development) {
    & $VenvPython -c "import docling, pytest, reportlab"
    if ($LASTEXITCODE -ne 0) {
        throw "Windows 开发依赖导入失败。"
    }
}

if ($Development) {
    Write-Host "Windows 开发与生产环境已就绪：$VenvPython"
}
else {
    Write-Host "Windows 生产环境已就绪：$VenvPython"
}
Write-Host "运行生产页面真值前，请确认 WPS Writer 已安装；脚本会通过只读 COM 打开原题。"
