[CmdletBinding()]
param(
    [switch]$RunTests
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"

if (-not (Test-Path $VenvPython)) {
    throw "未找到项目虚拟环境。请先运行 .\scripts\setup_windows.ps1 -Development。"
}

& $VenvPython -c "import pdfplumber, pypdfium2, PIL, docx, pyautogui, win32com.client"
if ($LASTEXITCODE -ne 0) {
    throw "Windows 运行依赖导入失败。"
}

$WpsProgIds = @("KWPS.Application", "wps.Application")
$RegisteredWps = $WpsProgIds | Where-Object {
    Test-Path "Registry::HKEY_CLASSES_ROOT\$_\CLSID"
} | Select-Object -First 1

if (-not $RegisteredWps) {
    throw "未检测到 WPS Writer COM 注册（KWPS.Application / wps.Application）。Microsoft Word 不作为回退。"
}

Write-Host "WPS COM 注册检查通过：$RegisteredWps"

if ($RunTests) {
    & $VenvPython -m pytest -q -p no:cacheprovider `
        test_document_render.py `
        test_macos_quicklook_render.py `
        test_document_visual_review.py `
        test_document_family_calibration.py `
        test_p1b_cli.py `
        test_p1b_batch.py `
        test_document_preflight.py `
        test_document_families.py `
        test_windows_scripts.py
    if ($LASTEXITCODE -ne 0) {
        throw "P1b 离线测试失败。"
    }
}

Write-Host "Windows 环境自检通过；未启动 WPS、未打开文档、未执行任何按键。"
