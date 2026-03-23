# ============================================
# SETUP COMPLETO - PC NOVO
# Coleta info | BitLocker | Usuario Admin
# Instalacao de programas
# Compativel com Windows 10 e 11
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
$arquivo = "$desktop\info-pc.txt"
$data = Get-Date -Format "dd/MM/yyyy HH:mm"
$etapaTotal = 7
$erros = @()

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  SETUP COMPLETO - PC NOVO" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan

# ============================================
# [1] COLETA DE INFORMACOES
# ============================================

Write-Host "`n[1/$etapaTotal] Coletando informacoes do PC..." -ForegroundColor Cyan

$serial = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber
$cpu = (Get-CimInstance -ClassName Win32_Processor).Name
$ramBytes = (Get-CimInstance -ClassName Win32_ComputerSystem).TotalPhysicalMemory
$ramGB = [math]::Round($ramBytes / 1GB, 1)
$nomePC = $env:COMPUTERNAME
$modelo = (Get-CimInstance -ClassName Win32_ComputerSystem).Model
$winVer = (Get-CimInstance -ClassName Win32_OperatingSystem).Caption

$discos = Get-PhysicalDisk | Select-Object MediaType, FriendlyName, @{
    Name = "Tamanho (GB)";
    Expression = { [math]::Round($_.Size / 1GB, 0) }
}

Write-Host "  OK" -ForegroundColor Green

# ============================================
# [2] REMOVER BLOATWARE
# ============================================

Write-Host "`n[2/$etapaTotal] Removendo bloatware..." -ForegroundColor Cyan

$bloatware = @(
    # McAfee
    "McAfee*",
    # Microsoft
    "Microsoft.OneDrive*",
    "Microsoft.MicrosoftTeams*",
    "MicrosoftTeams*",
    "Microsoft.Todos*",
    "Microsoft.MicrosoftSolitaireCollection*",
    "Microsoft.MicrosoftOfficeHub*",
    "Microsoft.BingNews*",
    "Microsoft.BingWeather*",
    "Microsoft.GetHelp*",
    "Microsoft.Getstarted*",
    "Microsoft.WindowsMail*",
    "Microsoft.windowscommunicationsapps*",
    "microsoft.windowscomm*",
    "Microsoft.SkypeApp*",
    "Microsoft.LinkedIn*",
    "Microsoft.Clipchamp*",
    "Microsoft.GamingApp*",
    "Microsoft.XboxApp*",
    "Microsoft.XboxGameOverlay*",
    "Microsoft.XboxGamingOverlay*",
    "Microsoft.XboxSpeechToTextOverlay*",
    "Microsoft.XboxIdentityProvider*",
    "Microsoft.Xbox.TCUI*",
    # Terceiros
    "SpotifyAB.SpotifyMusic*",
    "king.com.CandyCrushSaga*",
    "king.com.CandyCrush*",
    "BytedancePte.Ltd.TikTok*",
    "Facebook*",
    "Instagram*",
    "Disney*",
    "Clipchamp*"
)

$removidos = @()

foreach ($app in $bloatware) {
    $pacotes = Get-AppxPackage -AllUsers -Name $app -ErrorAction SilentlyContinue
    foreach ($pacote in $pacotes) {
        try {
            Remove-AppxPackage -Package $pacote.PackageFullName -AllUsers -ErrorAction Stop
            $removidos += $pacote.Name
            Write-Host "  Removido: $($pacote.Name)" -ForegroundColor Green
        } catch {
            # Tenta via provisioned package
            Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName -like $app } |
                ForEach-Object {
                    Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction SilentlyContinue
                    $removidos += $_.DisplayName
                    Write-Host "  Removido (provisioned): $($_.DisplayName)" -ForegroundColor Green
                }
        }
    }

    # Remover provisioned packages para nao voltar em novos usuarios
    Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like $app } |
        ForEach-Object {
            Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction SilentlyContinue
        }
}

# Remover McAfee via WMI/registry (nao e app da Store)
$mcafee = Get-WmiObject -Class Win32_Product -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*McAfee*" }
foreach ($m in $mcafee) {
    $m.Uninstall() | Out-Null
    $removidos += $m.Name
    Write-Host "  Removido: $($m.Name)" -ForegroundColor Green
}

# Desinstalar OneDrive completamente
$onedrivePath = "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
if (-not (Test-Path $onedrivePath)) { $onedrivePath = "$env:SystemRoot\System32\OneDriveSetup.exe" }
if (Test-Path $onedrivePath) {
    Start-Process $onedrivePath -ArgumentList "/uninstall" -Wait -ErrorAction SilentlyContinue
    Write-Host "  OneDrive desinstalado" -ForegroundColor Green
    $removidos += "OneDrive"
}

if ($removidos.Count -eq 0) {
    Write-Host "  Nenhum bloatware encontrado" -ForegroundColor Gray
} else {
    Write-Host "  $($removidos.Count) programa(s) removido(s)" -ForegroundColor Green
}

# ============================================
# [3] DESATIVAR BITLOCKER
# ============================================

Write-Host "`n[3/$etapaTotal] Verificando BitLocker..." -ForegroundColor Cyan
$bitlockerStatus = ""
try {
    $volumes = Get-BitLockerVolume -ErrorAction Stop
    foreach ($vol in $volumes) {
        if ($vol.ProtectionStatus -eq "On") {
            Disable-BitLocker -MountPoint $vol.MountPoint -ErrorAction Stop
            $bitlockerStatus += "  $($vol.MountPoint) - DESATIVADO`n"
            Write-Host "  BitLocker desativado em $($vol.MountPoint)" -ForegroundColor Green
        } else {
            $bitlockerStatus += "  $($vol.MountPoint) - Ja desativado`n"
            Write-Host "  $($vol.MountPoint) ja estava sem BitLocker" -ForegroundColor Gray
        }
    }
} catch {
    $bitlockerStatus = "  Nao disponivel ou nao ativo"
    Write-Host "  BitLocker nao encontrado" -ForegroundColor Gray
}

# ============================================
# [4] CRIAR USUARIO ADMIN
# ============================================

Write-Host "`n[4/$etapaTotal] Criando usuario Admin..." -ForegroundColor Cyan
$usuarioExiste = Get-LocalUser -Name "Admin" -ErrorAction SilentlyContinue
if ($usuarioExiste) {
    Write-Host "  Usuario 'Admin' ja existe" -ForegroundColor Yellow
    $adminStatus = "Ja existia"
} else {
    try {
        $senha = ConvertTo-SecureString "1010" -AsPlainText -Force
        New-LocalUser -Name "Admin" -Password $senha -FullName "Admin" -Description "Conta de manutencao" -PasswordNeverExpires -ErrorAction Stop
        $adminGroup = (Get-LocalGroup | Where-Object { $_.SID -like "S-1-5-32-544" }).Name
        Add-LocalGroupMember -Group $adminGroup -Member "Admin" -ErrorAction SilentlyContinue
        Write-Host "  Usuario 'Admin' criado!" -ForegroundColor Green
        $adminStatus = "Criado com sucesso"
    } catch {
        Write-Host "  Erro: $_" -ForegroundColor Red
        $adminStatus = "Erro: $_"
        $erros += "Usuario Admin"
    }
}

# ============================================
# [5] INSTALAR PROGRAMAS (ultima versao)
# ============================================

Write-Host "`n[5/$etapaTotal] Instalando programas..." -ForegroundColor Cyan

$programas = @(
    @{ nome = "Google Chrome";   id = "Google.Chrome" },
    @{ nome = "KeePass 2";      id = "DominikReichl.KeePass" },
    @{ nome = "WinRAR";         id = "RARLab.WinRAR" },
    @{ nome = "AnyDesk";        id = "AnyDeskSoftware.AnyDesk" },
    @{ nome = "Slack";          id = "SlackTechnologies.Slack" }
)

$instalados = @()
$total = $programas.Count
$atual = 0

foreach ($prog in $programas) {
    $atual++
    Write-Host "  [$atual/$total] $($prog.nome)..." -ForegroundColor Yellow -NoNewline

    $resultado = winget install --id $prog.id -e --accept-source-agreements --accept-package-agreements --silent 2>&1

    if ($LASTEXITCODE -eq 0) {
        Write-Host " OK" -ForegroundColor Green
        $instalados += "$($prog.nome) - Instalado"
    } elseif ($resultado -match "already installed") {
        Write-Host " Ja instalado" -ForegroundColor Gray
        $instalados += "$($prog.nome) - Ja instalado"
    } else {
        Write-Host " ERRO" -ForegroundColor Red
        $instalados += "$($prog.nome) - ERRO"
        $erros += $prog.nome
    }
}

# ============================================
# [6] INSTALAR OPENVPN (versao estavel)
# ============================================

Write-Host "`n[6/$etapaTotal] Instalando OpenVPN..." -ForegroundColor Cyan

$resultado = winget install --id OpenVPNTechnologies.OpenVPN -e --accept-source-agreements --accept-package-agreements --silent 2>&1

if ($LASTEXITCODE -eq 0) {
    Write-Host "  OpenVPN instalado!" -ForegroundColor Green
    $instalados += "OpenVPN - Instalado"
} elseif ($resultado -match "already installed") {
    Write-Host "  OpenVPN ja instalado" -ForegroundColor Gray
    $instalados += "OpenVPN - Ja instalado"
} else {
    Write-Host "  ERRO ao instalar OpenVPN" -ForegroundColor Red
    $instalados += "OpenVPN - ERRO"
    $erros += "OpenVPN"
}

# ============================================
# [7] GERAR RELATORIO
# ============================================

Write-Host "`n[7/$etapaTotal] Gerando relatorio..." -ForegroundColor Cyan

$conteudo = @"
==========================================
  INFORMACOES DO PC - $data
==========================================

Nome do PC    : $nomePC
Modelo        : $modelo
Serial Number : $serial
Processador   : $cpu
RAM           : $ramGB GB
Windows       : $winVer

------------------------------------------
  DISCO(S) DE ARMAZENAMENTO
------------------------------------------
"@

foreach ($disco in $discos) {
    $conteudo += "`n  $($disco.FriendlyName) - $($disco.'Tamanho (GB)') GB ($($disco.MediaType))"
}

$conteudo += @"

------------------------------------------
  CONFIGURACOES
------------------------------------------
Bloatware removido: $($removidos.Count) programa(s)

BitLocker:
$bitlockerStatus
Usuario Admin : $adminStatus

------------------------------------------
  PROGRAMAS
------------------------------------------
"@

foreach ($inst in $instalados) {
    $conteudo += "`n  $inst"
}

if ($erros.Count -gt 0) {
    $conteudo += "`n`n------------------------------------------"
    $conteudo += "`n  ERROS"
    $conteudo += "`n------------------------------------------"
    foreach ($e in $erros) {
        $conteudo += "`n  ! $e"
    }
}

$conteudo += "`n`n=========================================="

$conteudo | Out-File -FilePath $arquivo -Encoding UTF8

# ============================================
# RESUMO FINAL
# ============================================

Write-Host ""
Write-Host "=========================================" -ForegroundColor Green
Write-Host "  SETUP CONCLUIDO!" -ForegroundColor Green
Write-Host "  Relatorio: $arquivo" -ForegroundColor Green
if ($erros.Count -gt 0) {
    Write-Host "  Erros: $($erros.Count) programa(s) falharam" -ForegroundColor Red
} else {
    Write-Host "  Tudo instalado sem erros!" -ForegroundColor Green
}
Write-Host "=========================================" -ForegroundColor Green
Write-Host ""
Write-Host $conteudo
pause
