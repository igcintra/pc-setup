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

# Corrigir DNS para evitar falha de resolucao de nomes
Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | ForEach-Object {
    Set-DnsClientServerAddress -InterfaceIndex $_.ifIndex -ServerAddresses ("8.8.8.8","8.8.4.4") -ErrorAction SilentlyContinue
}
Write-Host "DNS configurado (Google 8.8.8.8)" -ForegroundColor Gray

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
    "MSTeams*",
    "Microsoft.Teams*",
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
Stop-Process -Name "OneDrive" -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
$onedrivePath = "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
if (-not (Test-Path $onedrivePath)) { $onedrivePath = "$env:SystemRoot\System32\OneDriveSetup.exe" }
if (Test-Path $onedrivePath) {
    Start-Process $onedrivePath -ArgumentList "/uninstall" -Wait -ErrorAction SilentlyContinue
    Write-Host "  OneDrive desinstalado" -ForegroundColor Green
    $removidos += "OneDrive"
}
# Remover via winget tambem
winget uninstall --id Microsoft.OneDrive -e --silent 2>&1 | Out-Null

# Desinstalar Teams completamente (Win 10 e 11)
Stop-Process -Name "ms-teams" -Force -ErrorAction SilentlyContinue
Stop-Process -Name "Teams" -Force -ErrorAction SilentlyContinue
winget uninstall --id Microsoft.Teams -e --silent 2>&1 | Out-Null
winget uninstall --name "Microsoft Teams" --silent 2>&1 | Out-Null
# Teams classico
$teamsPath = "$env:LOCALAPPDATA\Microsoft\Teams\Update.exe"
if (Test-Path $teamsPath) {
    Start-Process $teamsPath -ArgumentList "--uninstall -s" -Wait -ErrorAction SilentlyContinue
    Write-Host "  Teams desinstalado" -ForegroundColor Green
    $removidos += "Teams"
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
    @{ nome = "AnyDesk";        id = "AnyDeskSoftware.AnyDesk"; fallback = $true },
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
        # Fallback: download direto para AnyDesk
        if ($prog.fallback) {
            Write-Host " Winget falhou, baixando direto..." -ForegroundColor Yellow
            try {
                $anydeskMsi = "https://download.anydesk.com/AnyDesk.msi"
                $anydeskInstaller = "$env:TEMP\AnyDesk.msi"
                Invoke-WebRequest -Uri $anydeskMsi -OutFile $anydeskInstaller -ErrorAction Stop
                Start-Process msiexec.exe -ArgumentList "/i `"$anydeskInstaller`" /qn /norestart" -Wait -ErrorAction SilentlyContinue
                Remove-Item $anydeskInstaller -Force -ErrorAction SilentlyContinue
                Write-Host " OK (download direto)" -ForegroundColor Green
                $instalados += "$($prog.nome) - Instalado (download direto)"
            } catch {
                Write-Host " ERRO" -ForegroundColor Red
                $instalados += "$($prog.nome) - ERRO"
                $erros += $prog.nome
            }
        } else {
            Write-Host " ERRO" -ForegroundColor Red
            $instalados += "$($prog.nome) - ERRO"
            $erros += $prog.nome
        }
    }
}

# Definir Chrome como navegador padrao
$chromePath = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
if (-not (Test-Path $chromePath)) { $chromePath = "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe" }
if (Test-Path $chromePath) {
    Write-Host "  Configurando Chrome como navegador padrao..." -ForegroundColor Yellow
    $regBase = "HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations"
    $protocols = @("http", "https")
    foreach ($proto in $protocols) {
        $regPath = "$regBase\$proto\UserChoice"
        Remove-Item -Path $regPath -Force -ErrorAction SilentlyContinue
        New-Item -Path $regPath -Force -ErrorAction SilentlyContinue | Out-Null
        Set-ItemProperty -Path $regPath -Name "ProgId" -Value "ChromeHTML" -ErrorAction SilentlyContinue
    }
    # Associar extensoes de arquivo
    $fileTypes = @(".htm", ".html", ".shtml", ".xhtml")
    foreach ($ext in $fileTypes) {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$ext\UserChoice"
        Remove-Item -Path $regPath -Force -ErrorAction SilentlyContinue
        New-Item -Path $regPath -Force -ErrorAction SilentlyContinue | Out-Null
        Set-ItemProperty -Path $regPath -Name "ProgId" -Value "ChromeHTML" -ErrorAction SilentlyContinue
    }
    Write-Host "  Chrome definido como navegador padrao" -ForegroundColor Green
}

# Criar atalho do KeePass 2 na Area de Trabalho
$keepassExe = "${env:ProgramFiles(x86)}\KeePass Password Safe 2\KeePass.exe"
if (-not (Test-Path $keepassExe)) { $keepassExe = "$env:ProgramFiles\KeePass Password Safe 2\KeePass.exe" }
if (Test-Path $keepassExe) {
    $shell = New-Object -ComObject WScript.Shell
    $atalho = $shell.CreateShortcut("$desktop\KeePass 2.lnk")
    $atalho.TargetPath = $keepassExe
    $atalho.WorkingDirectory = (Split-Path $keepassExe)
    $atalho.Save()
    Write-Host "  Atalho KeePass 2 criado na Area de Trabalho" -ForegroundColor Green
}

# ============================================
# [6] INSTALAR OPENVPN 2.4.7 (versao fixa)
# ============================================

Write-Host "`n[6/$etapaTotal] Instalando OpenVPN 2.4.7..." -ForegroundColor Cyan

# INSTRUCAO: Substitua o link abaixo pelo link do Google Drive com o instalador
$openvpnUrl = "https://drive.google.com/uc?export=download&id=1H_i2cSJJGKT4lLfqD5HYHk9sJ-k-mC01"
$openvpnInstaller = "$env:TEMP\openvpn-install-2.4.7.exe"

try {
    Invoke-WebRequest -Uri $openvpnUrl -OutFile $openvpnInstaller -ErrorAction Stop
    Start-Process $openvpnInstaller -ArgumentList "/S" -Wait -ErrorAction SilentlyContinue
    Remove-Item $openvpnInstaller -Force -ErrorAction SilentlyContinue
    Write-Host "  OpenVPN 2.4.7 instalado!" -ForegroundColor Green
    $instalados += "OpenVPN 2.4.7 - Instalado"
} catch {
    Write-Host "  ERRO ao instalar OpenVPN" -ForegroundColor Red
    $instalados += "OpenVPN 2.4.7 - ERRO"
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
