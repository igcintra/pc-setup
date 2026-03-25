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

# Contador de uso
Invoke-RestMethod -Uri "https://script.google.com/macros/s/AKfycbwZwJrHL2SnECPzx5inz2K5_AVxbVvukXMra0grAgSbVuNjbxeNnP8sLDGdy-Sf2yfvoA/exec?script=pc-setup" -ErrorAction SilentlyContinue | Out-Null

$desktop = [Environment]::GetFolderPath("Desktop")
$arquivo = "$desktop\info-pc.txt"
$data = Get-Date -Format "dd/MM/yyyy HH:mm"
$etapaTotal = 12
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
Stop-Process -Name "OneDriveSetup" -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
$onedrivePath = "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
if (-not (Test-Path $onedrivePath)) { $onedrivePath = "$env:SystemRoot\System32\OneDriveSetup.exe" }
if (Test-Path $onedrivePath) {
    $proc = Start-Process $onedrivePath -ArgumentList "/uninstall" -PassThru -ErrorAction SilentlyContinue
    if ($proc -and -not $proc.WaitForExit(30000)) {
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        Write-Host "  OneDrive: timeout, forcado" -ForegroundColor Yellow
    } else {
        Write-Host "  OneDrive desinstalado" -ForegroundColor Green
    }
    $removidos += "OneDrive"
}
# Remover via winget tambem (timeout 30s)
$wingetOD = Start-Process "winget" -ArgumentList "uninstall --id Microsoft.OneDrive -e --silent" -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
if ($wingetOD -and -not $wingetOD.WaitForExit(30000)) {
    Stop-Process -Id $wingetOD.Id -Force -ErrorAction SilentlyContinue
}

# Desinstalar Teams completamente (Win 10 e 11)
Stop-Process -Name "ms-teams" -Force -ErrorAction SilentlyContinue
Stop-Process -Name "Teams" -Force -ErrorAction SilentlyContinue
$wingetTeams1 = Start-Process "winget" -ArgumentList "uninstall --id Microsoft.Teams -e --silent" -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
if ($wingetTeams1 -and -not $wingetTeams1.WaitForExit(30000)) {
    Stop-Process -Id $wingetTeams1.Id -Force -ErrorAction SilentlyContinue
}
$wingetTeams2 = Start-Process "winget" -ArgumentList "uninstall --name `"Microsoft Teams`" --silent" -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
if ($wingetTeams2 -and -not $wingetTeams2.WaitForExit(30000)) {
    Stop-Process -Id $wingetTeams2.Id -Force -ErrorAction SilentlyContinue
}
# Teams classico
$teamsPath = "$env:LOCALAPPDATA\Microsoft\Teams\Update.exe"
if (Test-Path $teamsPath) {
    $proc = Start-Process $teamsPath -ArgumentList "--uninstall -s" -PassThru -ErrorAction SilentlyContinue
    if ($proc -and -not $proc.WaitForExit(30000)) {
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
    }
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
    @{ nome = "Slack";          id = "SlackTechnologies.Slack" }
)

$instalados = @()
$total = $programas.Count + 1  # +1 para AnyDesk separado
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

# AnyDesk - baixa e abre instalador (precisa clicar em Instalar)
$atual++
Write-Host "  [$atual/$total] AnyDesk..." -ForegroundColor Yellow -NoNewline
$anydeskServico = Get-Service -Name "AnyDesk" -ErrorAction SilentlyContinue
if ($anydeskServico) {
    Write-Host " Ja instalado" -ForegroundColor Gray
    $instalados += "AnyDesk - Ja instalado"
} else {
    try {
        $anydeskUrl = "https://github.com/igcintra/pc-setup/releases/download/v1.0/AnyDesk.exe"
        $anydeskTemp = "$env:TEMP\AnyDesk.exe"
        Invoke-WebRequest -Uri $anydeskUrl -OutFile $anydeskTemp -ErrorAction Stop
        Write-Host ""
        Write-Host ""
        Write-Host "  ============================================" -ForegroundColor Yellow
        Write-Host "  ATENCAO: AnyDesk vai abrir." -ForegroundColor Yellow
        Write-Host "  Clique em 'Instalar AnyDesk' no programa." -ForegroundColor Yellow
        Write-Host "  Depois feche a janela do AnyDesk." -ForegroundColor Yellow
        Write-Host "  O script continua automaticamente." -ForegroundColor Yellow
        Write-Host "  ============================================" -ForegroundColor Yellow
        Write-Host ""
        Start-Process $anydeskTemp -Wait
        Remove-Item $anydeskTemp -Force -ErrorAction SilentlyContinue
        # Verificar se instalou como servico
        Start-Sleep -Seconds 3
        $anydeskServico = Get-Service -Name "AnyDesk" -ErrorAction SilentlyContinue
        if ($anydeskServico) {
            Write-Host "  AnyDesk instalado com sucesso!" -ForegroundColor Green
            $instalados += "AnyDesk - Instalado"
        } else {
            Write-Host "  AnyDesk pode nao ter sido instalado corretamente" -ForegroundColor Yellow
            $instalados += "AnyDesk - Verificar manualmente"
        }
    } catch {
        Write-Host " ERRO: $_" -ForegroundColor Red
        $instalados += "AnyDesk - ERRO"
        $erros += "AnyDesk"
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

$openvpnUrl = "https://github.com/igcintra/pc-setup/releases/download/v1.0/openvpn-install-2.4.7-I607-Win10.exe"
$openvpnInstaller = "$env:TEMP\openvpn-install-2.4.7.exe"

try {
    Invoke-WebRequest -Uri $openvpnUrl -OutFile $openvpnInstaller -ErrorAction Stop
    Start-Process $openvpnInstaller -ArgumentList "/S" -Wait -ErrorAction SilentlyContinue
    Remove-Item $openvpnInstaller -Force -ErrorAction SilentlyContinue
    Write-Host "  OpenVPN 2.4.7 instalado!" -ForegroundColor Green
    $instalados += "OpenVPN 2.4.7 - Instalado"

    # Ajustar TODOS os atalhos do OpenVPN: iniciar em config + executar como admin
    $openvpnShortcuts = @(
        "$env:PUBLIC\Desktop\OpenVPN GUI.lnk",
        "$desktop\OpenVPN GUI.lnk",
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\OpenVPN\OpenVPN GUI.lnk"
    )
    foreach ($lnk in $openvpnShortcuts) {
        if (Test-Path $lnk) {
            # Mudar "Iniciar em" para config
            $shell = New-Object -ComObject WScript.Shell
            $atalho = $shell.CreateShortcut($lnk)
            $atalho.WorkingDirectory = "$env:ProgramFiles\OpenVPN\config"
            $atalho.Save()
            # Forcar executar como administrador (flag byte no .lnk)
            $bytes = [System.IO.File]::ReadAllBytes($lnk)
            $bytes[0x15] = $bytes[0x15] -bor 0x20
            [System.IO.File]::WriteAllBytes($lnk, $bytes)
            Write-Host "  Atalho ajustado: $lnk" -ForegroundColor Green
        }
    }
    Write-Host "  OpenVPN: config + executar como admin" -ForegroundColor Green
} catch {
    Write-Host "  ERRO ao instalar OpenVPN: $_" -ForegroundColor Red
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

# ============================================
# [8] DESATIVAR NOTIFICACOES DO WINDOWS
# ============================================

Write-Host "`n[8/$etapaTotal] Desativando notificacoes..." -ForegroundColor Cyan

try {
    $regNotif = "HKCU:\Software\Microsoft\Windows\CurrentVersion\PushNotifications"
    if (-not (Test-Path $regNotif)) { New-Item -Path $regNotif -Force | Out-Null }
    Set-ItemProperty -Path $regNotif -Name "ToastEnabled" -Value 0 -Type DWord

    $regAction = "HKCU:\Software\Policies\Microsoft\Windows\Explorer"
    if (-not (Test-Path $regAction)) { New-Item -Path $regAction -Force | Out-Null }
    Set-ItemProperty -Path $regAction -Name "DisableNotificationCenter" -Value 1 -Type DWord

    $regLock = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings"
    if (-not (Test-Path $regLock)) { New-Item -Path $regLock -Force | Out-Null }
    Set-ItemProperty -Path $regLock -Name "NOC_GLOBAL_SETTING_ALLOW_NOTIFICATION_SOUND" -Value 0 -Type DWord
    Set-ItemProperty -Path $regLock -Name "NOC_GLOBAL_SETTING_ALLOW_TOASTS_ABOVE_LOCK" -Value 0 -Type DWord

    $regSugest = "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
    if (Test-Path $regSugest) {
        Set-ItemProperty -Path $regSugest -Name "SubscribedContent-338389Enabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $regSugest -Name "SubscribedContent-310093Enabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $regSugest -Name "SubscribedContent-338388Enabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $regSugest -Name "SoftLandingEnabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
    }

    $regWPN = "HKCU:\Software\Microsoft\Windows\CurrentVersion\PushNotifications"
    Set-ItemProperty -Path $regWPN -Name "DatabaseMigrationCompleted" -Value 1 -Type DWord -ErrorAction SilentlyContinue

    $regNotifSettings = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings"
    if (-not (Test-Path $regNotifSettings)) { New-Item -Path $regNotifSettings -Force | Out-Null }
    Set-ItemProperty -Path $regNotifSettings -Name "NOC_GLOBAL_SETTING_TOASTS_ENABLED" -Value 0 -Type DWord

    $regPolicy = "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\PushNotifications"
    if (-not (Test-Path $regPolicy)) { New-Item -Path $regPolicy -Force | Out-Null }
    Set-ItemProperty -Path $regPolicy -Name "NoToastApplicationNotification" -Value 1 -Type DWord

    $regTips = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
    if (-not (Test-Path $regTips)) { New-Item -Path $regTips -Force | Out-Null }
    Set-ItemProperty -Path $regTips -Name "DisableSoftLanding" -Value 1 -Type DWord -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $regTips -Name "DisableWindowsConsumerFeatures" -Value 1 -Type DWord -ErrorAction SilentlyContinue

    Write-Host "  Todas as notificacoes desativadas" -ForegroundColor Green
} catch {
    Write-Host "  ERRO" -ForegroundColor Red
}

# ============================================
# [9] LIMPAR BARRA DE TAREFAS E FIXAR PROGRAMAS
# ============================================

Write-Host "`n[9/$etapaTotal] Configurando barra de tarefas..." -ForegroundColor Cyan

try {
    $regAdvanced = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Set-ItemProperty -Path $regAdvanced -Name "ShowTaskViewButton" -Value 0 -Type DWord -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $regAdvanced -Name "TaskbarDa" -Value 0 -Type DWord -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $regAdvanced -Name "TaskbarMn" -Value 0 -Type DWord -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $regAdvanced -Name "ShowCortanaButton" -Value 0 -Type DWord -ErrorAction SilentlyContinue

    $regSearch = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
    if (-not (Test-Path $regSearch)) { New-Item -Path $regSearch -Force | Out-Null }
    Set-ItemProperty -Path $regSearch -Name "SearchboxTaskbarMode" -Value 0 -Type DWord

    $regCortana = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
    if (-not (Test-Path $regCortana)) { New-Item -Path $regCortana -Force | Out-Null }
    Set-ItemProperty -Path $regCortana -Name "AllowCortana" -Value 0 -Type DWord -ErrorAction SilentlyContinue

    # Remover Noticias e Interesses / Tempo (Win 10)
    try {
        $regFeeds = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Feeds"
        New-Item -Path $regFeeds -Force -ErrorAction SilentlyContinue | Out-Null
        New-ItemProperty -Path $regFeeds -Name "ShellFeedsTaskbarViewMode" -Value 2 -PropertyType DWord -Force -ErrorAction SilentlyContinue | Out-Null
        New-ItemProperty -Path $regFeeds -Name "IsFeedsAvailable" -Value 0 -PropertyType DWord -Force -ErrorAction SilentlyContinue | Out-Null
    } catch { }

    # Remover Widgets (Win 11) via politica
    $regWidgets = "HKLM:\SOFTWARE\Policies\Microsoft\Dsh"
    if (-not (Test-Path $regWidgets)) { New-Item -Path $regWidgets -Force | Out-Null }
    Set-ItemProperty -Path $regWidgets -Name "AllowNewsAndInterests" -Value 0 -Type DWord -ErrorAction SilentlyContinue

    # ---- LIMPAR MENU INICIAR (tiles/pins) ----

    # Win 10: Remover todos os tiles do Menu Iniciar
    $startTiles = (New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}')
    if ($startTiles) {
        $startTiles.Items() | ForEach-Object {
            $_.Verbs() | Where-Object { $_.Name -match "Unpin|Desafixar|Desanclar" } | ForEach-Object { $_.DoIt() }
        }
    }

    # Win 10/11: Limpar cache de tiles do registro
    $startCachePath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\CloudStore\Store\Cache\DefaultAccount"
    if (Test-Path $startCachePath) {
        Get-ChildItem $startCachePath -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match "start\.tilegrid" } |
            ForEach-Object { Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue }
    }

    # Win 11: Limpar layout do Menu Iniciar (remover todos os pins)
    $startLayoutPath = "$env:LOCALAPPDATA\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState"
    if (Test-Path $startLayoutPath) {
        Remove-Item "$startLayoutPath\start*.bin" -Force -ErrorAction SilentlyContinue
        Remove-Item "$startLayoutPath\start2.bin" -Force -ErrorAction SilentlyContinue
    }

    Write-Host "  Menu Iniciar limpo" -ForegroundColor Green

    $pinDir = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
    if (Test-Path $pinDir) { Remove-Item "$pinDir\*" -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Path $pinDir -Force | Out-Null

    $taskbandPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband"
    Remove-Item -Path $taskbandPath -Force -Recurse -ErrorAction SilentlyContinue
    New-Item -Path $taskbandPath -Force | Out-Null

    # Desafixar Microsoft Edge e Microsoft Store da barra
    $regPins = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband\AuxilliaryPins"
    Remove-Item -Path $regPins -Force -Recurse -ErrorAction SilentlyContinue

    # Explorador de Arquivos (unico item fixado)
    $shell = New-Object -ComObject WScript.Shell
    $atalho = $shell.CreateShortcut("$pinDir\01-Explorador de Arquivos.lnk")
    $atalho.TargetPath = "explorer.exe"
    $atalho.Save()

    Write-Host "  Barra limpa (apenas Explorador de Arquivos)" -ForegroundColor Green
} catch {
    Write-Host "  ERRO" -ForegroundColor Red
}

# ============================================
# [10] WALLPAPER IG NETWORKS
# ============================================

Write-Host "`n[10/$etapaTotal] Configurando wallpaper..." -ForegroundColor Cyan

try {
    $wpUrl = "https://github.com/igcintra/pc-setup/releases/download/v1.0/IGN.jpg"
    $wpPath = "$env:USERPROFILE\Pictures\IGN-wallpaper.jpg"
    Invoke-WebRequest -Uri $wpUrl -OutFile $wpPath -UseBasicParsing -ErrorAction Stop

    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class Wallpaper {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
    }
"@
    [Wallpaper]::SystemParametersInfo(0x0014, 0, $wpPath, 0x0003) | Out-Null

    # Estilo: Fill (preencher)
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WallpaperStyle" -Value "10" -ErrorAction SilentlyContinue
    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "TileWallpaper" -Value "0" -ErrorAction SilentlyContinue

    Write-Host "  Wallpaper IG Networks aplicado" -ForegroundColor Green
} catch {
    Write-Host "  ERRO ao aplicar wallpaper: $_" -ForegroundColor Red
    $erros += "Wallpaper"
}

# ============================================
# [11] REMOVER AUTO-INICIO DE PROGRAMAS
# ============================================

Write-Host "`n[11/$etapaTotal] Removendo programas do inicio automatico..." -ForegroundColor Cyan

# Itens que DEVEM permanecer no auto-inicio
$manter = @("SecurityHealth", "RtkAudUService")

# Limpar HKCU Run (remover TUDO exceto os mantidos e AnyDesk)
$regRun = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
if (Test-Path $regRun) {
    $entries = Get-ItemProperty $regRun -ErrorAction SilentlyContinue
    foreach ($prop in $entries.PSObject.Properties) {
        if ($prop.Name -match "^PS" -or $prop.Name -eq "(default)") { continue }
        $keep = $false
        foreach ($m in $manter) { if ($prop.Name -like "*$m*") { $keep = $true } }
        if ($prop.Name -like "*AnyDesk*") { $keep = $true }
        if (-not $keep) {
            Remove-ItemProperty -Path $regRun -Name $prop.Name -ErrorAction SilentlyContinue
            Write-Host "  Removido HKCU: $($prop.Name)" -ForegroundColor Green
        }
    }
}

# Limpar HKLM Run (remover TUDO exceto mantidos e AnyDesk)
$regRunLM = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"
if (Test-Path $regRunLM) {
    $entries = Get-ItemProperty $regRunLM -ErrorAction SilentlyContinue
    foreach ($prop in $entries.PSObject.Properties) {
        if ($prop.Name -match "^PS" -or $prop.Name -eq "(default)") { continue }
        $keep = $false
        foreach ($m in $manter) { if ($prop.Name -like "*$m*") { $keep = $true } }
        if ($prop.Name -like "*AnyDesk*") { $keep = $true }
        if (-not $keep) {
            Remove-ItemProperty -Path $regRunLM -Name $prop.Name -ErrorAction SilentlyContinue
            Write-Host "  Removido HKLM: $($prop.Name)" -ForegroundColor Green
        }
    }
}

# Corrigir AnyDesk para executar em segundo plano (--control)
$anydeskPaths = @(
    "$env:ProgramFiles\AnyDesk\AnyDesk.exe",
    "${env:ProgramFiles(x86)}\AnyDesk\AnyDesk.exe"
)
foreach ($adPath in $anydeskPaths) {
    if (Test-Path $adPath) {
        Set-ItemProperty -Path $regRunLM -Name "AnyDesk" -Value "`"$adPath`" --control" -ErrorAction SilentlyContinue
        Write-Host "  AnyDesk configurado em segundo plano (--control)" -ForegroundColor Green
        break
    }
}

# Limpar pasta Startup do usuario (tudo)
$startupFolder = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
if (Test-Path $startupFolder) {
    Get-ChildItem $startupFolder -ErrorAction SilentlyContinue | ForEach-Object {
        Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
        Write-Host "  Removido Startup: $($_.Name)" -ForegroundColor Green
    }
}

# Limpar pasta Common Startup (todos os usuarios)
$commonStartup = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
if (Test-Path $commonStartup) {
    Get-ChildItem $commonStartup -ErrorAction SilentlyContinue | ForEach-Object {
        Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
        Write-Host "  Removido Common Startup: $($_.Name)" -ForegroundColor Green
    }
}

# Desativar auto-inicio via Task Manager (StartupApproved)
$permitidos = @("SecurityHealth", "RtkAudUService", "AnyDesk")
$disabledBytes = [byte[]](0x03,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00)

# StartupApproved HKCU
$regApproved = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
if (Test-Path $regApproved) {
    (Get-Item $regApproved).GetValueNames() | ForEach-Object {
        if ($_ -eq "(default)") { return }
        $permitido = $false
        foreach ($p in $permitidos) { if ($_ -like "*$p*") { $permitido = $true } }
        if (-not $permitido) {
            Set-ItemProperty -Path $regApproved -Name $_ -Value $disabledBytes -Type Binary -ErrorAction SilentlyContinue
            Write-Host "  Desativado startup: $_" -ForegroundColor Green
        }
    }
}

# StartupApproved HKLM
$regApprovedLM = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
if (Test-Path $regApprovedLM) {
    (Get-Item $regApprovedLM).GetValueNames() | ForEach-Object {
        if ($_ -eq "(default)") { return }
        $permitido = $false
        foreach ($p in $permitidos) { if ($_ -like "*$p*") { $permitido = $true } }
        if (-not $permitido) {
            Set-ItemProperty -Path $regApprovedLM -Name $_ -Value $disabledBytes -Type Binary -ErrorAction SilentlyContinue
            Write-Host "  Desativado startup HKLM: $_" -ForegroundColor Green
        }
    }
}

# StartupApproved\StartupFolder
$regApprovedFolder = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder"
if (Test-Path $regApprovedFolder) {
    (Get-Item $regApprovedFolder).GetValueNames() | ForEach-Object {
        if ($_ -eq "(default)") { return }
        $permitido = $false
        foreach ($p in $permitidos) { if ($_ -like "*$p*") { $permitido = $true } }
        if (-not $permitido) {
            Set-ItemProperty -Path $regApprovedFolder -Name $_ -Value $disabledBytes -Type Binary -ErrorAction SilentlyContinue
            Write-Host "  Desativado startup folder: $_" -ForegroundColor Green
        }
    }
}

# Remover OneDrive do auto-inicio (persistente)
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "OneDrive" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "OneDriveSetup" -ErrorAction SilentlyContinue

# Remover programas especificos que se readicionam
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "Discord" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "com.squirrel.slack.slack" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "Steam" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "EpicGamesLauncher" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "LGHUB" -ErrorAction SilentlyContinue

# Desativar TODAS as tarefas agendadas de logon (exceto do sistema)
$tarefasManter = @("MicrosoftEdgeUpdateTask", "SecurityHealth", "Windows", "Microsoft\Windows")
Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
    $_.Triggers | Where-Object { $_ -is [Microsoft.Management.Infrastructure.CimInstance] -and $_.CimClass.CimClassName -eq "MSFT_TaskLogonTrigger" }
} | ForEach-Object {
    $skip = $false
    foreach ($m in $tarefasManter) { if ($_.TaskPath -like "*$m*") { $skip = $true } }
    if (-not $skip) {
        Disable-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -ErrorAction SilentlyContinue
        Write-Host "  Tarefa desativada: $($_.TaskName)" -ForegroundColor Green
    }
}

Write-Host "  Auto-inicio limpo (AnyDesk segundo plano + audio mantidos)" -ForegroundColor Green

# ============================================
# [12] VERIFICAR CONTA MICROSOFT E CONVERTER PARA LOCAL
# ============================================

Write-Host "`n[12/$etapaTotal] Verificando conta do usuario..." -ForegroundColor Cyan

$usuarioLogado = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName
if ($usuarioLogado) {
    $nomeUsuario = $usuarioLogado.Split("\")[-1]
    Write-Host "  Usuario logado: $usuarioLogado" -ForegroundColor Gray

    $userInfo = Get-LocalUser -Name $nomeUsuario -ErrorAction SilentlyContinue
    $isMicrosoftAccount = $false

    if ($userInfo -and $userInfo.PrincipalSource -eq "MicrosoftAccount") {
        $isMicrosoftAccount = $true
    }

    # Verificar tambem pelo SID
    $profileList = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    Get-ChildItem $profileList -ErrorAction SilentlyContinue | ForEach-Object {
        $profilePath = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).ProfileImagePath
        if ($profilePath -and $profilePath -match $nomeUsuario) {
            if ($_.PSChildName -match "^S-1-12-") {
                $isMicrosoftAccount = $true
            }
        }
    }

    if ($isMicrosoftAccount) {
        Write-Host ""
        Write-Host "  ============================================" -ForegroundColor Yellow
        Write-Host "  ATENCAO: '$nomeUsuario' esta vinculado" -ForegroundColor Yellow
        Write-Host "  a uma conta Microsoft!" -ForegroundColor Yellow
        Write-Host "  ============================================" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Abrindo configuracoes para converter..." -ForegroundColor Yellow
        Write-Host "  Va em: Suas informacoes > Entrar com" -ForegroundColor Yellow
        Write-Host "  conta local" -ForegroundColor Yellow
        Write-Host ""
        Start-Process "ms-settings:yourinfo"
    } else {
        Write-Host "  '$nomeUsuario' ja e conta local" -ForegroundColor Green
    }
} else {
    Write-Host "  Nao foi possivel identificar o usuario" -ForegroundColor Yellow
}

pause
