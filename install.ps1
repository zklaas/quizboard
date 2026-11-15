# Install Quizboard on Windows (current user, no admin needed).
#
#   powershell -ExecutionPolicy Bypass -File install.ps1
#
# Creates .venv in this folder with PyQt6, writes the `quizboard` command into a
# per-user bin folder on PATH, and adds a Start-menu shortcut with the Quizboard icon.
param([string]$BinDir = "")

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $BinDir) { $BinDir = Join-Path $env:LOCALAPPDATA 'Programs\Quizboard\bin' }

function Get-PyVersion([string]$Exe) {
    if (-not $Exe -or -not (Test-Path -LiteralPath $Exe)) { return $null }
    if ($Exe -match '\\WindowsApps\\') { return $null }       # Microsoft Store stub
    try {
        # no inner double quotes: Windows PowerShell 5.1 strips them from native args
        $v = & $Exe -c 'import sys; print(*sys.version_info[:2], sep=chr(46))' 2>$null
        if ($LASTEXITCODE -eq 0) { return "$v".Trim() }
    } catch { }
    return $null
}

function Test-PyOk([string]$Ver) {
    if (-not $Ver) { return $false }
    $p = $Ver.Split('.'); $n = [int]$p[0] * 100 + [int]$p[1]
    return ($n -ge 310)
}

function Find-Python {
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) {
        foreach ($v in @('3.13', '3.12', '3.14', '3.11', '3.10')) {
            try {
                $exe = (& $py.Source "-$v" -c 'import sys; print(sys.executable)' 2>$null)
                if ($exe -and (Test-PyOk (Get-PyVersion $exe.Trim()))) { return $exe.Trim() }
            } catch { }
        }
    }
    foreach ($v in @('313', '312', '314', '311', '310')) {
        $exe = Join-Path $env:LOCALAPPDATA ("Programs\Python\Python" + $v + "\python.exe")
        if (Test-PyOk (Get-PyVersion $exe)) { return $exe }
    }
    $onPath = Get-Command python -ErrorAction SilentlyContinue
    if ($onPath -and (Test-PyOk (Get-PyVersion $onPath.Source))) { return $onPath.Source }
    throw "Quizboard needs Python 3.10 or newer.`nInstall it, then run this again:`n    winget install --id Python.Python.3.13 -e --scope user"
}

$python = Find-Python
Write-Host "Python: $python"
$venv = Join-Path $AppDir '.venv'
$vpy = Join-Path $venv 'Scripts\python.exe'
$vpyw = Join-Path $venv 'Scripts\pythonw.exe'
if ((Test-Path -LiteralPath $vpy) -and -not (Test-PyOk (Get-PyVersion $vpy))) {
    Write-Host "Replacing .venv (it was made with another Python version)."
    Remove-Item -Recurse -Force $venv
}
if (-not (Test-Path -LiteralPath $vpy)) {
    & $python -m venv $venv
    if ($LASTEXITCODE -ne 0) { throw "Could not create $venv" }
}
& $vpy -m pip install --quiet --upgrade pip
& $vpy -m pip install --quiet PyQt6
if ($LASTEXITCODE -ne 0) { throw "pip install failed" }

# `quizboard` command
New-Item -ItemType Directory -Force -Path $BinDir | Out-Null
$shim = Join-Path $BinDir 'quizboard.cmd'
Set-Content -LiteralPath $shim -Encoding ASCII -Value "@`"$vpy`" `"$AppDir\quizboard.py`" %*"
Write-Host "Installed: $shim"
$userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
if (-not $userPath) { $userPath = '' }
if (($userPath -split ';') -notcontains $BinDir) {
    [Environment]::SetEnvironmentVariable('Path', ($userPath.TrimEnd(';') + ';' + $BinDir).TrimStart(';'), 'User')
    Write-Host "Added $BinDir to your PATH (open a new terminal to use it)."
}

# Start-menu shortcut (pythonw: no console window)
$startMenu = Join-Path ([Environment]::GetFolderPath('Programs')) 'Quizboard.lnk'
$ws = New-Object -ComObject WScript.Shell
$lnk = $ws.CreateShortcut($startMenu)
$lnk.TargetPath = $vpyw
$lnk.Arguments = "`"$AppDir\quizboard.py`""
$lnk.WorkingDirectory = $AppDir
$lnk.IconLocation = (Join-Path $AppDir 'icons\quizboard.ico')
$lnk.Description = "A 6 x 6 quiz board read from an Excel workbook"
$lnk.Save()
Write-Host "Start menu: $startMenu"
Write-Host "Done. Run it with:  quizboard      (or from the Start menu)"
