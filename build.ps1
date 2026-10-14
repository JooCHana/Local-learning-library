$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $ProjectDir

$VenvPython = Join-Path $ProjectDir ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $VenvPython)) {
    python -m venv .venv
}

& $VenvPython -m pip install --disable-pip-version-check `
    -i "https://pypi.tuna.tsinghua.edu.cn/simple" -r requirements.txt
if ($LASTEXITCODE -ne 0) {
    Write-Host "镜像源暂不可用，改用官方软件源重试…"
    & $VenvPython -m pip install --disable-pip-version-check `
        --index-url "https://pypi.org/simple" -r requirements.txt
}
if ($LASTEXITCODE -ne 0) {
    throw "依赖安装失败，退出代码：$LASTEXITCODE"
}
& $VenvPython -m PyInstaller --noconfirm --clean --onefile --windowed `
    --name "智能文档学习库3.6.4" `
    --icon "assets\app_icon.ico" `
    --add-data "assets\app_icon.ico;assets" `
    --hidden-import win32com.client `
    --hidden-import pythoncom `
    main.py

if ($LASTEXITCODE -ne 0) {
    throw "程序打包失败，退出代码：$LASTEXITCODE"
}
Write-Host "生成完成：$ProjectDir\dist\智能文档学习库3.6.4.exe"
