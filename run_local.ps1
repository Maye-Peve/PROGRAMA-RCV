param(
  [string]$WorkbookPath,
  [int]$Port = 8765
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$bundledPython = Join-Path ([Environment]::GetFolderPath('UserProfile')) '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
$pythonCandidates = @()

if (-not $WorkbookPath) {
  $userProfile = [Environment]::GetFolderPath('UserProfile')
  $WorkbookPath = Join-Path (Join-Path $userProfile 'Downloads') 'RCVSEP2026 (1).xlsx'
}

if ($env:RCV_PYTHON) {
  $pythonCandidates += $env:RCV_PYTHON
}

$pythonCommand = Get-Command python -ErrorAction SilentlyContinue
if ($pythonCommand) {
  $pythonCandidates += $pythonCommand.Source
}

$pyCommand = Get-Command py -ErrorAction SilentlyContinue
if ($pyCommand) {
  $pythonCandidates += $pyCommand.Source
}

if (Test-Path -LiteralPath $bundledPython) {
  $pythonCandidates += $bundledPython
}

$pythonExe = $null
$pythonArgs = @()
foreach ($candidate in ($pythonCandidates | Select-Object -Unique)) {
  try {
    if ($pyCommand -and $candidate -eq $pyCommand.Source) {
      & $candidate -3 -c 'import openpyxl' 2>$null
      if ($LASTEXITCODE -eq 0) {
        $pythonExe = $candidate
        $pythonArgs = @('-3')
        break
      }
    } else {
      & $candidate -c 'import openpyxl' 2>$null
      if ($LASTEXITCODE -eq 0) {
        $pythonExe = $candidate
        break
      }
    }
  } catch {
    continue
  }
}

if (-not $pythonExe) {
  Write-Error 'No se encontró Python con openpyxl. Instala Python y ejecuta: python -m pip install -r requirements.txt'
}

if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
  Write-Error "No se encontró el libro: $WorkbookPath. Pasa otra ruta con -WorkbookPath."
}

Set-Location -LiteralPath $projectRoot
Write-Host 'Iniciando la aplicación local. Para detenerla, presiona Ctrl+C.'
& $pythonExe @pythonArgs (Join-Path $projectRoot 'server.py') --workbook $WorkbookPath --port $Port
