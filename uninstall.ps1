# ============================================
# DESINSTALAR PROGRAMAS DO SETUP
# Script para testes - remove tudo que o setup instala
# ============================================

# Verificar se esta rodando como Administrador
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "ERRO: Execute este script como Administrador!" -ForegroundColor Red
    Write-Host "Clique com botao direito no PowerShell > Executar como Administrador" -ForegroundColor Yellow
    pause
    exit
}

Write-Host ""
Write-Host "=========================================" -ForegroundColor Red
Write-Host "  DESINSTALACAO - MODO TESTE" -ForegroundColor Red
Write-Host "=========================================" -ForegroundColor Red
Write-Host ""

# ============================================
# PROGRAMAS VIA WINGET
# ============================================

$programas = @(
    @{ nome = "Google Chrome";   id = "Google.Chrome" },
    @{ nome = "KeePass 2";      id = "DominikReichl.KeePass" },
    @{ nome = "WinRAR";         id = "RARLab.WinRAR" },
    @{ nome = "Slack";          id = "SlackTechnologies.Slack" }
)

$total = $programas.Count + 2  # +2 para AnyDesk e OpenVPN
$atual = 0

foreach ($prog in $programas) {
    $atual++
    Write-Host "[$atual/$total] Removendo $($prog.nome)..." -ForegroundColor Yellow -NoNewline

    $resultado = winget uninstall --id $prog.id -e --silent 2>&1

    if ($LASTEXITCODE -eq 0) {
        Write-Host " OK" -ForegroundColor Green
    } else {
        Write-Host " Ja desinstalado" -ForegroundColor Gray
    }
}

# ============================================
# ANYDESK - Remocao direta
# ============================================

$atual++
Write-Host "[$atual/$total] Removendo AnyDesk..." -ForegroundColor Yellow

$anydeskRemovido = $false

# Fechar processo
Stop-Process -Name "AnyDesk" -Force -ErrorAction SilentlyContinue

# Metodo 1: Desinstalador proprio
$anydeskPaths = @(
    "${env:ProgramFiles(x86)}\AnyDesk\AnyDesk.exe",
    "$env:ProgramFiles\AnyDesk\AnyDesk.exe"
)
foreach ($path in $anydeskPaths) {
    if (Test-Path $path) {
        Start-Process $path -ArgumentList "--remove --silent" -Wait -ErrorAction SilentlyContinue
        Write-Host "  OK - Removido via desinstalador" -ForegroundColor Green
        $anydeskRemovido = $true
        break
    }
}

# Metodo 2: Via registro
if (-not $anydeskRemovido) {
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($regPath in $regPaths) {
        $entry = Get-ItemProperty $regPath -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like "*AnyDesk*" }
        if ($entry -and $entry.UninstallString) {
            Start-Process cmd.exe -ArgumentList "/c `"$($entry.UninstallString)`" --silent" -Wait -ErrorAction SilentlyContinue
            Write-Host "  OK - Removido via registro" -ForegroundColor Green
            $anydeskRemovido = $true
            break
        }
    }
}

# Metodo 3: Winget
if (-not $anydeskRemovido) {
    winget uninstall --id AnyDeskSoftware.AnyDesk -e --silent 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  OK - Removido via winget" -ForegroundColor Green
        $anydeskRemovido = $true
    }
}

if (-not $anydeskRemovido) {
    Write-Host "  Ja desinstalado" -ForegroundColor Gray
}

# Limpar pasta residual
Remove-Item "${env:ProgramFiles(x86)}\AnyDesk" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$env:ProgramFiles\AnyDesk" -Recurse -Force -ErrorAction SilentlyContinue

# ============================================
# OPENVPN - Remocao direta
# ============================================

$atual++
Write-Host "[$atual/$total] Removendo OpenVPN..." -ForegroundColor Yellow

$openvpnRemovido = $false

# Fechar processos
Stop-Process -Name "openvpn*" -Force -ErrorAction SilentlyContinue
Stop-Process -Name "openvpnserv*" -Force -ErrorAction SilentlyContinue

# Metodo 1: Desinstalador proprio
$openvpnUninstall = "$env:ProgramFiles\OpenVPN\Uninstall.exe"
if (Test-Path $openvpnUninstall) {
    Start-Process $openvpnUninstall -ArgumentList "/S" -Wait -ErrorAction SilentlyContinue
    Write-Host "  OK - Removido via desinstalador" -ForegroundColor Green
    $openvpnRemovido = $true
}

# Metodo 2: Via registro (MSI)
if (-not $openvpnRemovido) {
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($regPath in $regPaths) {
        $entry = Get-ItemProperty $regPath -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like "*OpenVPN*" }
        if ($entry) {
            if ($entry.UninstallString -match "msiexec") {
                $productCode = $entry.PSChildName
                Start-Process msiexec.exe -ArgumentList "/x $productCode /qn /norestart" -Wait -ErrorAction SilentlyContinue
            } else {
                Start-Process cmd.exe -ArgumentList "/c `"$($entry.UninstallString)`" /S" -Wait -ErrorAction SilentlyContinue
            }
            Write-Host "  OK - Removido via registro" -ForegroundColor Green
            $openvpnRemovido = $true
            break
        }
    }
}

# Metodo 3: Winget
if (-not $openvpnRemovido) {
    winget uninstall --id OpenVPNTechnologies.OpenVPN -e --silent 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  OK - Removido via winget" -ForegroundColor Green
        $openvpnRemovido = $true
    }
}

if (-not $openvpnRemovido) {
    Write-Host "  Ja desinstalado" -ForegroundColor Gray
}

# Limpar pasta residual
Remove-Item "$env:ProgramFiles\OpenVPN" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "=========================================" -ForegroundColor Green
Write-Host "  DESINSTALACAO CONCLUIDA!" -ForegroundColor Green
Write-Host "  Agora pode rodar o setup.ps1 do zero" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green
pause
