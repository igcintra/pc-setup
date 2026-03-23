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

$programas = @(
    @{ nome = "Google Chrome";   id = "Google.Chrome" },
    @{ nome = "KeePass 2";      id = "DominikReichl.KeePass" },
    @{ nome = "WinRAR";         id = "RARLab.WinRAR" },
    @{ nome = "AnyDesk";        id = "AnyDeskSoftware.AnyDesk" },
    @{ nome = "Slack";          id = "SlackTechnologies.Slack" },
    @{ nome = "OpenVPN";        id = "OpenVPNTechnologies.OpenVPN" }
)

$total = $programas.Count
$atual = 0

foreach ($prog in $programas) {
    $atual++
    Write-Host "[$atual/$total] Removendo $($prog.nome)..." -ForegroundColor Yellow -NoNewline

    $resultado = winget uninstall --id $prog.id -e --silent 2>&1

    if ($LASTEXITCODE -eq 0) {
        Write-Host " OK" -ForegroundColor Green
    } elseif ($resultado -match "No installed package") {
        Write-Host " Nao encontrado" -ForegroundColor Gray
    } else {
        Write-Host " Winget falhou, tentando metodo alternativo..." -ForegroundColor Yellow
        $removido = $false

        # Tentar via registro do Windows (uninstall string)
        $regPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        foreach ($regPath in $regPaths) {
            $entry = Get-ItemProperty $regPath -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName -like "*$($prog.nome)*" }
            if ($entry) {
                $uninstallCmd = $entry.UninstallString
                if ($entry.QuietUninstallString) { $uninstallCmd = $entry.QuietUninstallString }
                if ($uninstallCmd) {
                    Start-Process cmd.exe -ArgumentList "/c $uninstallCmd /S /silent /quiet /norestart" -Wait -ErrorAction SilentlyContinue
                    Write-Host "  Removido via registro" -ForegroundColor Green
                    $removido = $true
                    break
                }
            }
        }

        if (-not $removido) {
            Write-Host "  Nao foi possivel remover automaticamente" -ForegroundColor Red
        }
    }
}

Write-Host ""
Write-Host "=========================================" -ForegroundColor Green
Write-Host "  DESINSTALACAO CONCLUIDA!" -ForegroundColor Green
Write-Host "  Agora pode rodar o setup.ps1 do zero" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green
pause
