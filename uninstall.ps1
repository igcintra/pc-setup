# ============================================
# DESINSTALAR PROGRAMAS DO SETUP
# Script para testes - remove tudo que o setup instala
# Gera log detalhado na Area de Trabalho
# ============================================

# Verificar se esta rodando como Administrador
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "ERRO: Execute este script como Administrador!" -ForegroundColor Red
    Write-Host "Clique com botao direito no PowerShell > Executar como Administrador" -ForegroundColor Yellow
    pause
    exit
}

$desktop = [Environment]::GetFolderPath("Desktop")
$logFile = "$desktop\uninstall-log.txt"
$log = @("========== LOG DE DESINSTALACAO ==========", "Data: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')", "")

function Log($msg) {
    $script:log += $msg
    Write-Host $msg
}

Write-Host ""
Write-Host "=========================================" -ForegroundColor Red
Write-Host "  DESINSTALACAO - MODO TESTE" -ForegroundColor Red
Write-Host "=========================================" -ForegroundColor Red
Write-Host ""

# ============================================
# LISTAR TUDO QUE O WINGET VE INSTALADO
# ============================================

Log "--- PROGRAMAS DETECTADOS PELO WINGET ---"
$wingetList = winget list 2>&1
$log += $wingetList | Out-String
Log ""

Log "--- PROGRAMAS NO REGISTRO DO WINDOWS ---"
$regPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)
foreach ($regPath in $regPaths) {
    Get-ItemProperty $regPath -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like "*AnyDesk*" -or $_.DisplayName -like "*OpenVPN*" } |
        ForEach-Object {
            Log "  Nome: $($_.DisplayName)"
            Log "  UninstallString: $($_.UninstallString)"
            Log "  QuietUninstallString: $($_.QuietUninstallString)"
            Log "  InstallLocation: $($_.InstallLocation)"
            Log "  RegPath: $($_.PSPath)"
            Log ""
        }
}

# ============================================
# PROGRAMAS VIA WINGET
# ============================================

Log "--- REMOVENDO PROGRAMAS VIA WINGET ---"

$programas = @(
    @{ nome = "Google Chrome";   id = "Google.Chrome" },
    @{ nome = "KeePass 2";      id = "DominikReichl.KeePass" },
    @{ nome = "Slack";          id = "SlackTechnologies.Slack" }
)

$total = $programas.Count + 3  # +3 para WinRAR, AnyDesk e OpenVPN
$atual = 0

foreach ($prog in $programas) {
    $atual++
    Log "[$atual/$total] Removendo $($prog.nome)..."

    $resultado = winget uninstall --id $prog.id -e --silent 2>&1
    $exitCode = $LASTEXITCODE

    Log "  Exit code: $exitCode"
    Log "  Resultado: $($resultado | Out-String)"

    if ($exitCode -eq 0) {
        Log "  STATUS: OK"
    } else {
        Log "  STATUS: Ja desinstalado ou nao encontrado"
    }
    Log ""
}

# ============================================
# WINRAR - Remocao direta (nao respeita --silent do winget)
# ============================================

$atual++
Log "[$atual/$total] Removendo WinRAR..."
$winrarRemovido = $false

$winrarUninstall = "$env:ProgramFiles\WinRAR\uninstall.exe"
if (Test-Path $winrarUninstall) {
    Log "  Executando: $winrarUninstall /S"
    $proc = Start-Process $winrarUninstall -ArgumentList "/S" -Wait -PassThru -ErrorAction SilentlyContinue
    Log "  Exit code: $($proc.ExitCode)"
    Log "  OK - Removido via desinstalador"
    $winrarRemovido = $true
}

if (-not $winrarRemovido) {
    # Tentar x86
    $winrarUninstall = "${env:ProgramFiles(x86)}\WinRAR\uninstall.exe"
    if (Test-Path $winrarUninstall) {
        $proc = Start-Process $winrarUninstall -ArgumentList "/S" -Wait -PassThru -ErrorAction SilentlyContinue
        Log "  OK - Removido via desinstalador (x86)"
        $winrarRemovido = $true
    }
}

if (-not $winrarRemovido) {
    Log "  Ja desinstalado"
}

Remove-Item "$env:ProgramFiles\WinRAR" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "${env:ProgramFiles(x86)}\WinRAR" -Recurse -Force -ErrorAction SilentlyContinue
Log ""

# ============================================
# ANYDESK
# ============================================

$atual++
Log "[$atual/$total] Removendo AnyDesk..."
$anydeskRemovido = $false

Stop-Process -Name "AnyDesk" -Force -ErrorAction SilentlyContinue
Log "  Processo AnyDesk finalizado (se existia)"

# Verificar onde esta instalado
$anydeskPaths = @(
    "${env:ProgramFiles(x86)}\AnyDesk\AnyDesk.exe",
    "$env:ProgramFiles\AnyDesk\AnyDesk.exe",
    "$env:APPDATA\AnyDesk\AnyDesk.exe",
    "$env:LOCALAPPDATA\AnyDesk\AnyDesk.exe"
)

foreach ($path in $anydeskPaths) {
    $existe = Test-Path $path
    Log "  Verificando $path : $existe"
    if ($existe -and -not $anydeskRemovido) {
        Log "  Executando: $path --remove --silent"
        $proc = Start-Process $path -ArgumentList "--remove --silent" -Wait -PassThru -ErrorAction SilentlyContinue
        Log "  Exit code: $($proc.ExitCode)"
        $anydeskRemovido = $true
    }
}

# Via registro
if (-not $anydeskRemovido) {
    Log "  Tentando via registro..."
    foreach ($regPath in $regPaths) {
        $entries = Get-ItemProperty $regPath -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like "*AnyDesk*" }
        foreach ($entry in $entries) {
            Log "  Encontrado: $($entry.DisplayName)"
            Log "  UninstallString: $($entry.UninstallString)"
            Log "  QuietUninstallString: $($entry.QuietUninstallString)"
            $cmd = if ($entry.QuietUninstallString) { $entry.QuietUninstallString } else { "$($entry.UninstallString) --silent" }
            Log "  Executando: cmd /c $cmd"
            $proc = Start-Process cmd.exe -ArgumentList "/c $cmd" -Wait -PassThru -ErrorAction SilentlyContinue
            Log "  Exit code: $($proc.ExitCode)"
            $anydeskRemovido = $true
        }
    }
}

# Via winget
if (-not $anydeskRemovido) {
    Log "  Tentando via winget..."
    $resultado = winget uninstall --id AnyDeskSoftware.AnyDesk -e --silent 2>&1
    Log "  Exit code: $LASTEXITCODE"
    Log "  Resultado: $($resultado | Out-String)"
    if ($LASTEXITCODE -eq 0) { $anydeskRemovido = $true }
}

# Winget sem --silent
if (-not $anydeskRemovido) {
    Log "  Tentando winget sem --silent..."
    $resultado = winget uninstall --id AnyDeskSoftware.AnyDesk -e 2>&1
    Log "  Exit code: $LASTEXITCODE"
    Log "  Resultado: $($resultado | Out-String)"
    if ($LASTEXITCODE -eq 0) { $anydeskRemovido = $true }
}

# Winget por nome
if (-not $anydeskRemovido) {
    Log "  Tentando winget por nome..."
    $resultado = winget uninstall --name "AnyDesk" --silent 2>&1
    Log "  Exit code: $LASTEXITCODE"
    Log "  Resultado: $($resultado | Out-String)"
    if ($LASTEXITCODE -eq 0) { $anydeskRemovido = $true }
}

Log "  ANYDESK REMOVIDO: $anydeskRemovido"

Remove-Item "${env:ProgramFiles(x86)}\AnyDesk" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$env:ProgramFiles\AnyDesk" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$env:APPDATA\AnyDesk" -Recurse -Force -ErrorAction SilentlyContinue
Log ""

# ============================================
# OPENVPN
# ============================================

$atual++
Log "[$atual/$total] Removendo OpenVPN..."
$openvpnRemovido = $false

Stop-Process -Name "openvpn*" -Force -ErrorAction SilentlyContinue
Stop-Process -Name "openvpnserv*" -Force -ErrorAction SilentlyContinue
Log "  Processos OpenVPN finalizados (se existiam)"

# Verificar desinstalador
$openvpnUninstall = "$env:ProgramFiles\OpenVPN\Uninstall.exe"
$existe = Test-Path $openvpnUninstall
Log "  Verificando $openvpnUninstall : $existe"

if ($existe) {
    Log "  Executando: $openvpnUninstall /S"
    $proc = Start-Process $openvpnUninstall -ArgumentList "/S" -Wait -PassThru -ErrorAction SilentlyContinue
    Log "  Exit code: $($proc.ExitCode)"
    $openvpnRemovido = $true
}

# Via registro
if (-not $openvpnRemovido) {
    Log "  Tentando via registro..."
    foreach ($regPath in $regPaths) {
        $entries = Get-ItemProperty $regPath -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like "*OpenVPN*" }
        foreach ($entry in $entries) {
            Log "  Encontrado: $($entry.DisplayName)"
            Log "  UninstallString: $($entry.UninstallString)"
            if ($entry.UninstallString -match "msiexec") {
                $productCode = $entry.PSChildName
                Log "  Executando: msiexec /x $productCode /qn /norestart"
                $proc = Start-Process msiexec.exe -ArgumentList "/x $productCode /qn /norestart" -Wait -PassThru -ErrorAction SilentlyContinue
            } else {
                Log "  Executando: $($entry.UninstallString) /S"
                $proc = Start-Process cmd.exe -ArgumentList "/c `"$($entry.UninstallString)`" /S" -Wait -PassThru -ErrorAction SilentlyContinue
            }
            Log "  Exit code: $($proc.ExitCode)"
            $openvpnRemovido = $true
        }
    }
}

# Via winget
if (-not $openvpnRemovido) {
    Log "  Tentando via winget..."
    $resultado = winget uninstall --id OpenVPNTechnologies.OpenVPN -e --silent 2>&1
    Log "  Exit code: $LASTEXITCODE"
    Log "  Resultado: $($resultado | Out-String)"
    if ($LASTEXITCODE -eq 0) { $openvpnRemovido = $true }
}

# Winget sem --silent
if (-not $openvpnRemovido) {
    Log "  Tentando winget sem --silent..."
    $resultado = winget uninstall --id OpenVPNTechnologies.OpenVPN -e 2>&1
    Log "  Exit code: $LASTEXITCODE"
    Log "  Resultado: $($resultado | Out-String)"
    if ($LASTEXITCODE -eq 0) { $openvpnRemovido = $true }
}

# Winget por nome
if (-not $openvpnRemovido) {
    Log "  Tentando winget por nome..."
    $resultado = winget uninstall --name "OpenVPN" --silent 2>&1
    Log "  Exit code: $LASTEXITCODE"
    Log "  Resultado: $($resultado | Out-String)"
    if ($LASTEXITCODE -eq 0) { $openvpnRemovido = $true }
}

Log "  OPENVPN REMOVIDO: $openvpnRemovido"

Remove-Item "$env:ProgramFiles\OpenVPN" -Recurse -Force -ErrorAction SilentlyContinue
Log ""

# ============================================
# SALVAR LOG
# ============================================

$log += ""
$log += "========== FIM DO LOG =========="

$log | Out-File -FilePath $logFile -Encoding UTF8

Write-Host ""
Write-Host "=========================================" -ForegroundColor Green
Write-Host "  DESINSTALACAO CONCLUIDA!" -ForegroundColor Green
Write-Host "  Log salvo em: $logFile" -ForegroundColor Green
Write-Host "  Agora pode rodar o setup.ps1 do zero" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green
pause
