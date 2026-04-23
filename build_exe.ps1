Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

# Close running app (locks dist folder)
Get-Process -Name "RawMaterials" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

# Best-effort cleanup (sometimes antivirus/explorer keeps files locked)
for ($i = 0; $i -lt 5; $i++) {
  try {
    if (Test-Path ".\\dist\\RawMaterials") {
      Remove-Item -Recurse -Force ".\\dist\\RawMaterials"
    }
    break
  } catch {
    Start-Sleep -Milliseconds 700
  }
}

$extraData = @(
  "--add-data", "settings.json;."
)

if (Test-Path ".\\шаблон.docx") {
  $extraData += @("--add-data", "шаблон.docx;.")
} else {
  Write-Host "Warning: шаблон.docx not found рядом со скриптом; EXE соберется, но без шаблона."
}

pyinstaller --noconfirm --clean --windowed --name "RawMaterials" `
  @extraData `
  "app.py"

# Ensure settings.json is next to the EXE (PyInstaller may place data under _internal)
if (Test-Path ".\\settings.json") {
  Copy-Item -Force ".\\settings.json" ".\\dist\\RawMaterials\\settings.json"
}

Write-Host ""
Write-Host "Done. EXE location: dist\\RawMaterials\\RawMaterials.exe"

