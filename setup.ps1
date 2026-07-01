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

# Manter PC acordado durante toda a execucao do script
# (SetThreadExecutionState com ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED)
try {
    $sig = '[DllImport("kernel32.dll")] public static extern uint SetThreadExecutionState(uint esFlags);'
    $sleepHelper = Add-Type -MemberDefinition $sig -Name "Sleep" -Namespace "Win32" -PassThru -ErrorAction SilentlyContinue
    if ($sleepHelper) { $sleepHelper::SetThreadExecutionState(0x80000003) | Out-Null }
} catch {}

# ============================================
# Inicia log: tudo que aparece na tela vai pro arquivo
# Salva em Downloads do usuario com timestamp pra rastrear cada execucao
# ============================================
$logTimestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
$logPath = "$env:USERPROFILE\Downloads\pc-setup-log_$logTimestamp.txt"
try {
    Start-Transcript -Path $logPath -Force -ErrorAction Stop | Out-Null
    $logAtivo = $true
} catch { $logAtivo = $false }

Write-Host "========================================" -ForegroundColor DarkGray
Write-Host "  PC SETUP - LOG INICIADO" -ForegroundColor DarkGray
Write-Host "  $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor DarkGray
if ($logAtivo) { Write-Host "  Log: $logPath" -ForegroundColor DarkGray }
Write-Host "  PC: $env:COMPUTERNAME | User: $env:USERNAME" -ForegroundColor DarkGray
try {
    $os = Get-CimInstance Win32_OperatingSystem
    Write-Host "  OS: $($os.Caption) (build $($os.BuildNumber))" -ForegroundColor DarkGray
} catch {}
try {
    $cs = Get-CimInstance Win32_ComputerSystem
    Write-Host "  HW: $($cs.Manufacturer) $($cs.Model)" -ForegroundColor DarkGray
} catch {}
Write-Host "  PSVersion: $($PSVersionTable.PSVersion)" -ForegroundColor DarkGray
Write-Host "========================================" -ForegroundColor DarkGray
Write-Host ""

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
$etapaTotal = 13
$erros = @()

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  SETUP COMPLETO - PC NOVO" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan

# ============================================
# [1] COLETA DE INFORMACOES
# ============================================


# ==== CHECKPOINT / RETOMADA (add automatico) ====
$ckptDir  = "$env:ProgramData\pc-setup"
$ckptFile = "$ckptDir\state.txt"
function Get-Step {
    if (Test-Path $ckptFile) {
        $v = (Get-Content $ckptFile -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($v -match '^\d+$') { return [int]$v }
    }
    return 0
}
function Save-Step([int]$n) {
    try {
        if (-not (Test-Path $ckptDir)) { New-Item -ItemType Directory -Path $ckptDir -Force | Out-Null }
        Set-Content -Path $ckptFile -Value $n -Encoding ASCII
    } catch {}
}
$lastStep = Get-Step
if ($lastStep -gt 0) { Write-Host "`n>>> Retomando a partir da etapa $($lastStep + 1) (checkpoint encontrado em $ckptFile)." -ForegroundColor Magenta }
# defaults p/ o relatorio (etapa 7) nao quebrar se etapas foram puladas na retomada
$removidos = @(); $bitlockerStatus = "N/D"; $adminStatus = "N/D"; $instalados = @()
# ================================================
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

if ($lastStep -lt 2) {   # <<CKPT-OPEN 2>>
Write-Host "`n[2/$etapaTotal] Removendo bloatware..." -ForegroundColor Cyan

# Derrubar processos que podem interferir na remocao
$processosMatar = @(
    "McAfee*", "mcshield", "mcuicnt", "McUICnt", "ModuleCoreService", "MMSSHOST", "McPvTray",
    "OneDrive", "OneDriveSetup",
    "ms-teams", "Teams",
    "Dropbox", "DropboxUpdate",
    "WebAdvisor", "mcwebadvisor"
)
foreach ($proc in $processosMatar) {
    Get-Process -Name $proc -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}
Write-Host "  Processos bloatware encerrados" -ForegroundColor Green

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
    "Clipchamp*",
    # Dropbox
    "Dropbox*"
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

# McAfee: desativar servicos para nao reabrir, matar tudo
Write-Host "  McAfee: parando servicos..." -ForegroundColor Yellow
Get-Service -DisplayName "*McAfee*" -ErrorAction SilentlyContinue | ForEach-Object {
    Stop-Service -Name $_.Name -Force -ErrorAction SilentlyContinue
    Set-Service -Name $_.Name -StartupType Disabled -ErrorAction SilentlyContinue
}
function Kill-McAfee {
    Get-Process -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -match "mcafee|mcshield|mcuicnt|ModuleCore|MMSSHOST|McPvTray|WebAdvisor|McInstaller|mfemms|mfevtps|mcods|mfefire|mfetp|protectedmodulehost"
    } | Stop-Process -Force -ErrorAction SilentlyContinue
}
Kill-McAfee

$mcafeeIds = @("McAfee.WebAdvisor", "McAfee.McAfee", "McAfee.LiveSafe", "McAfee.TrueKey", "McAfee.SecurityScan")
foreach ($mid in $mcafeeIds) {
    Kill-McAfee
    Write-Host "  McAfee: $mid..." -ForegroundColor Yellow -NoNewline
    $p = Start-Process "winget" -ArgumentList "uninstall --id $mid -e --silent --force --disable-interactivity" -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
    if ($p) {
        if (-not $p.WaitForExit(15000)) {
            Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
            Kill-McAfee
            Write-Host " timeout (forcado)" -ForegroundColor Yellow
        } else {
            Write-Host " OK" -ForegroundColor Green
        }
    } else {
        Write-Host " nao encontrado" -ForegroundColor Gray
    }
}

# === Deep clean McAfee residual (WPS + WebAdvisor + tarefas + pastas) ===
# Cobre o que sobra em OEMs Lenovo/HP/Dell apos o winget uninstall
Write-Host "  McAfee: limpeza profunda silenciosa..." -ForegroundColor Yellow

# 1. Apagar tarefas agendadas (raiz das notificacoes do McAfee Anti-tracker, Health Check, etc)
$tasksRemoved = 0
Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like "*McAfee*" -or $_.TaskPath -like "*McAfee*" } | ForEach-Object {
    try {
        Unregister-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -Confirm:$false -ErrorAction Stop
        $tasksRemoved++
    } catch {}
}
if ($tasksRemoved -gt 0) { Write-Host "    $tasksRemoved tarefa(s) agendada(s) apagada(s)" -ForegroundColor DarkGray }

# 2. Rodar UninstallString do registro (silencioso, janela escondida)
$uninstKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)
$found = Get-ItemProperty $uninstKeys -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*McAfee*" }
foreach ($app in $found) {
    $unins = $app.QuietUninstallString
    if (-not $unins) { $unins = $app.UninstallString }
    if ($unins) {
        try {
            if ($unins -match '^"([^"]+)"\s*(.*)$') { $exe = $matches[1]; $argStr = $matches[2] }
            else { $tk = $unins -split ' ',2; $exe = $tk[0]; $argStr = if ($tk.Count -gt 1) { $tk[1] } else { '' } }
            if ($argStr -notmatch '/quiet|/qn|/silent|/S\b') { $argStr = "$argStr /quiet /norestart" }
            if (Test-Path $exe) {
                $p = Start-Process -FilePath $exe -ArgumentList $argStr -WindowStyle Hidden -PassThru -ErrorAction SilentlyContinue
                if ($p) {
                    if (-not $p.WaitForExit(180000)) { try { $p.Kill() } catch {} }
                    Write-Host "    desinstalado: $($app.DisplayName)" -ForegroundColor DarkGray
                }
            }
        } catch {}
    }
}

# 3. Apagar pastas residuais
$mcafeePaths = @("$env:ProgramFiles\McAfee","${env:ProgramFiles(x86)}\McAfee","$env:ProgramData\McAfee","$env:LOCALAPPDATA\McAfee","$env:APPDATA\McAfee")
foreach ($p in $mcafeePaths) {
    if (Test-Path $p) {
        Remove-Item $p -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# 4. Remover servico WebAdvisor residual via sc.exe
if (Get-Service -Name "McAfee WebAdvisor" -ErrorAction SilentlyContinue) {
    Start-Process sc.exe -ArgumentList "delete `"McAfee WebAdvisor`"" -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue | Out-Null
    Write-Host "    servico McAfee WebAdvisor removido" -ForegroundColor DarkGray
}

Write-Host "  McAfee: concluido" -ForegroundColor Green

# Matar TUDO via taskkill (nao trava em processos protegidos)
$matarNomes = @("OneDrive","OneDriveSetup","ms-teams","Teams","Dropbox","DropboxUpdate","mcafee*","McUICnt","mcshield","ModuleCoreService","MMSSHOST","McPvTray","WebAdvisor","mfemms","mfevtps","protectedmodulehost")
foreach ($nome in $matarNomes) {
    Start-Process "taskkill" -ArgumentList "/f /im $nome.exe" -WindowStyle Hidden -ErrorAction SilentlyContinue
}
Start-Sleep -Seconds 2

# OneDrive (timeout 15s)
Write-Host "  OneDrive..." -ForegroundColor Yellow -NoNewline
$onedrivePath = "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
if (-not (Test-Path $onedrivePath)) { $onedrivePath = "$env:SystemRoot\System32\OneDriveSetup.exe" }
if (Test-Path $onedrivePath) {
    $proc = Start-Process $onedrivePath -ArgumentList "/uninstall" -PassThru -ErrorAction SilentlyContinue
    if ($proc -and -not $proc.WaitForExit(15000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    $removidos += "OneDrive"
}
$p = Start-Process "winget" -ArgumentList "uninstall --id Microsoft.OneDrive -e --silent --force --disable-interactivity" -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
if ($p -and -not $p.WaitForExit(15000)) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
Write-Host " OK" -ForegroundColor Green

# Teams (timeout 15s)
Write-Host "  Teams..." -ForegroundColor Yellow -NoNewline
$p = Start-Process "winget" -ArgumentList "uninstall --id Microsoft.Teams -e --silent --force --disable-interactivity" -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
if ($p -and -not $p.WaitForExit(15000)) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
$teamsPath = "$env:LOCALAPPDATA\Microsoft\Teams\Update.exe"
if (Test-Path $teamsPath) {
    $proc = Start-Process $teamsPath -ArgumentList "--uninstall -s" -PassThru -ErrorAction SilentlyContinue
    if ($proc -and -not $proc.WaitForExit(15000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    $removidos += "Teams"
}
Write-Host " OK" -ForegroundColor Green

# Dropbox (timeout 15s)
Write-Host "  Dropbox..." -ForegroundColor Yellow -NoNewline
$p = Start-Process "winget" -ArgumentList "uninstall --id Dropbox.Dropbox -e --silent --force --disable-interactivity" -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
if ($p -and -not $p.WaitForExit(15000)) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
Write-Host " OK" -ForegroundColor Green

if ($removidos.Count -eq 0) {
    Write-Host "  Nenhum bloatware encontrado" -ForegroundColor Gray
} else {
    Write-Host "  $($removidos.Count) programa(s) removido(s)" -ForegroundColor Green
}

# ============================================
# [3] DESATIVAR BITLOCKER
# ============================================

Save-Step 2
}   # <<CKPT-CLOSE 2>>
if ($lastStep -lt 3) {   # <<CKPT-OPEN 3>>
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

Save-Step 3
}   # <<CKPT-CLOSE 3>>
if ($lastStep -lt 4) {   # <<CKPT-OPEN 4>>
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

Save-Step 4
}   # <<CKPT-CLOSE 4>>
if ($lastStep -lt 5) {   # <<CKPT-OPEN 5>>
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

# Habilita TLS 1.2 antes de qualquer download (necessario em Win10/11 fresh)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Helper de download com timeout REAL de transferencia (BITS, fallback curl.exe)
# Retorna $true se sucesso, $false se falha/timeout
function Download-File {
    param(
        [string]$Url,
        [string]$OutPath,
        [int]$TimeoutSec = 180
    )

    # Tentativa 1: BITS (servico nativo do Windows, com retry, resume, e timeout efetivo)
    try {
        $job = Start-BitsTransfer -Source $Url -Destination $OutPath -DisplayName "download-pc-setup" -Asynchronous -ErrorAction Stop
        $start = Get-Date
        while ($job.JobState -in 'Connecting','Transferring','Queued') {
            if (((Get-Date) - $start).TotalSeconds -gt $TimeoutSec) {
                Remove-BitsTransfer -BitsJob $job -ErrorAction SilentlyContinue
                Write-Host "    [download] TIMEOUT BITS apos $TimeoutSec s" -ForegroundColor Red
                return $false
            }
            Start-Sleep -Seconds 2
        }
        if ($job.JobState -eq 'Transferred') {
            Complete-BitsTransfer -BitsJob $job
            return (Test-Path $OutPath) -and ((Get-Item $OutPath).Length -gt 0)
        } else {
            Remove-BitsTransfer -BitsJob $job -ErrorAction SilentlyContinue
            Write-Host "    [download] BITS falhou (estado: $($job.JobState))" -ForegroundColor Red
        }
    } catch {
        Write-Host "    [download] BITS indisponivel, tentando curl.exe..." -ForegroundColor DarkYellow
    }

    # Tentativa 2: curl.exe (nativo Win10/11, --max-time funciona pra transferencia inteira)
    $curlExe = Get-Command curl.exe -ErrorAction SilentlyContinue
    if ($curlExe) {
        $p = Start-Process curl.exe -ArgumentList @("-L","-s","--max-time","$TimeoutSec","-o","`"$OutPath`"","`"$Url`"") -PassThru -Wait -NoNewWindow
        if ($p.ExitCode -eq 0 -and (Test-Path $OutPath) -and ((Get-Item $OutPath).Length -gt 0)) {
            return $true
        }
        Write-Host "    [download] curl.exe falhou (exit $($p.ExitCode))" -ForegroundColor Red
    }

    return $false
}

# ============================================
# Health check do winget + bootstrap se necessario
# ============================================
function Test-WingetWorking {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { return $false }
    try {
        $null = winget --version 2>&1
        if ($LASTEXITCODE -ne 0) { return $false }
        $sourceTest = winget source list 2>&1
        if ($LASTEXITCODE -ne 0) { return $false }
        return $true
    } catch { return $false }
}

function Install-WingetBootstrap {
    Write-Host "  Tentando bootstrap do winget (App Installer)..." -ForegroundColor Yellow
    $tempDir = "$env:TEMP\winget-bootstrap"
    if (-not (Test-Path $tempDir)) { New-Item -ItemType Directory -Path $tempDir -Force | Out-Null }

    try {
        # VCLibs (dependencia)
        $vcLibsUrl = "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx"
        Invoke-WebRequest -Uri $vcLibsUrl -OutFile "$tempDir\vclibs.appx" -UseBasicParsing -ErrorAction Stop

        # UI.Xaml (dependencia)
        $uiXamlUrl = "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx"
        Invoke-WebRequest -Uri $uiXamlUrl -OutFile "$tempDir\uixaml.appx" -UseBasicParsing -ErrorAction Stop

        # Winget (DesktopAppInstaller)
        $wingetUrl = "https://github.com/microsoft/winget-cli/releases/latest/download/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
        Invoke-WebRequest -Uri $wingetUrl -OutFile "$tempDir\winget.msixbundle" -UseBasicParsing -ErrorAction Stop

        Add-AppxPackage -Path "$tempDir\vclibs.appx" -ErrorAction SilentlyContinue
        Add-AppxPackage -Path "$tempDir\uixaml.appx" -ErrorAction SilentlyContinue
        Add-AppxPackage -Path "$tempDir\winget.msixbundle" -ErrorAction Stop

        Write-Host "  Bootstrap do winget concluido" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "  Bootstrap do winget falhou: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

if (-not (Test-WingetWorking)) {
    Write-Host "  Winget nao detectado ou nao funcional. Tentando bootstrap..." -ForegroundColor Yellow
    if (Install-WingetBootstrap) {
        # Recarrega PATH para que winget seja encontrado nesta mesma sessao
        $env:Path = [Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [Environment]::GetEnvironmentVariable("Path", "User")
        Start-Sleep -Seconds 3
        if (Test-WingetWorking) {
            winget source update 2>&1 | Out-Null
        }
    }
}

$wingetOk = Test-WingetWorking
if (-not $wingetOk) {
    Write-Host "  AVISO: Winget continua nao funcional. Programas via winget vao falhar." -ForegroundColor Red
    Write-Host "  Sugestao: instalar manualmente, ou atualizar 'App Installer' pela Microsoft Store." -ForegroundColor Yellow
}

# Desabilitar fonte msstore (causa erro 0x8a15005e em PCs com cert msstore vencido)
# Os programas que instalamos so existem na fonte winget, entao msstore nao perde nada
if ($wingetOk) {
    try {
        $msstoreCheck = winget source list 2>&1 | Out-String
        if ($msstoreCheck -match "msstore") {
            winget source remove msstore 2>&1 | Out-Null
            Write-Host "  Fonte msstore desabilitada (workaround pra cert vencido)" -ForegroundColor DarkGray
        }
    } catch {}
}

# Fallback de download direto (independente do winget) - por programa. Retorna $true se instalou.
function Install-Fallback {
    param([string]$Id)
    switch ($Id) {
        'Google.Chrome' {
            try {
                Write-Host "    [fallback] Baixando Chrome MSI direto do Google (timeout 4 min)..." -ForegroundColor DarkYellow
                $f = "$env:TEMP\chrome_installer.msi"
                if (Test-Path $f) { Remove-Item $f -Force -ErrorAction SilentlyContinue }
                if (Download-File "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi" $f 240) {
                    $p = Start-Process "msiexec.exe" -ArgumentList "/i `"$f`" /qn /norestart" -PassThru
                    if (-not $p.WaitForExit(180000)) { $p.Kill(); return $false }
                    $ok = ($p.ExitCode -eq 0); Remove-Item $f -Force -ErrorAction SilentlyContinue; return $ok
                }
            } catch {}
            return $false
        }
        'RARLab.WinRAR' {
            try {
                Write-Host "    [fallback] Baixando WinRAR direto do RARLab (timeout 2 min)..." -ForegroundColor DarkYellow
                $f = "$env:TEMP\winrar_installer.exe"
                if (Test-Path $f) { Remove-Item $f -Force -ErrorAction SilentlyContinue }
                if (Download-File "https://www.rarlab.com/rar/winrar-x64-711br.exe" $f 120) {
                    $p = Start-Process $f -ArgumentList "/S" -PassThru
                    if (-not $p.WaitForExit(120000)) { $p.Kill(); return $false }
                    $ok = ($p.ExitCode -eq 0); Remove-Item $f -Force -ErrorAction SilentlyContinue; return $ok
                }
            } catch {}
            return $false
        }
        'DominikReichl.KeePass' {
            try {
                Write-Host "    [fallback] Baixando KeePass 2 direto (SourceForge, timeout 3 min)..." -ForegroundColor DarkYellow
                $f = "$env:TEMP\keepass_setup.exe"
                if (Test-Path $f) { Remove-Item $f -Force -ErrorAction SilentlyContinue }
                if (Download-File "https://sourceforge.net/projects/keepass/files/KeePass%202.x/2.57.1/KeePass-2.57.1-Setup.exe/download" $f 180) {
                    $p = Start-Process $f -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART" -PassThru
                    if (-not $p.WaitForExit(120000)) { $p.Kill(); return $false }
                    $ok = ($p.ExitCode -eq 0); Remove-Item $f -Force -ErrorAction SilentlyContinue; return $ok
                }
            } catch {}
            return $false
        }
        'SlackTechnologies.Slack' {
            try {
                Write-Host "    [fallback] Baixando Slack MSI direto (timeout 3 min)..." -ForegroundColor DarkYellow
                $f = "$env:TEMP\slack_setup.msi"
                if (Test-Path $f) { Remove-Item $f -Force -ErrorAction SilentlyContinue }
                if (Download-File "https://slack.com/ssb/download-win64-msi" $f 180) {
                    $p = Start-Process "msiexec.exe" -ArgumentList "/i `"$f`" /qn /norestart" -PassThru
                    if (-not $p.WaitForExit(180000)) { $p.Kill(); return $false }
                    $ok = ($p.ExitCode -eq 0); Remove-Item $f -Force -ErrorAction SilentlyContinue; return $ok
                }
            } catch {}
            return $false
        }
    }
    return $false
}

foreach ($prog in $programas) {
    $atual++
    Write-Host "  [$atual/$total] $($prog.nome)..." -ForegroundColor Yellow -NoNewline

    $installed = $false
    $viaTxt = ""

    # 1) winget (se funcional)
    if ($wingetOk) {
        $resultado = winget install --id $prog.id -e --accept-source-agreements --accept-package-agreements --silent 2>&1
        $code = $LASTEXITCODE
        if ($code -eq 0) {
            $installed = $true; $viaTxt = "Instalado"
        } elseif ($resultado -match "already installed|ja esta instalado") {
            Write-Host " Ja instalado" -ForegroundColor Gray
            $instalados += "$($prog.nome) - Ja instalado"
            continue
        } else {
            $resultado2 = winget install --id $prog.id -e --source winget --accept-source-agreements --accept-package-agreements --silent 2>&1
            if ($LASTEXITCODE -eq 0) { $installed = $true; $viaTxt = "Instalado (retry)" }
        }
    }

    if ($installed) {
        Write-Host " OK" -ForegroundColor Green
        $instalados += "$($prog.nome) - $viaTxt"
        continue
    }

    # 2) fallback download direto (roda com winget indisponivel OU winget que falhou)
    Write-Host ""   # quebra a linha do NoNewline
    if (-not $wingetOk) {
        Write-Host "    [fallback] winget indisponivel -> tentando download direto..." -ForegroundColor DarkYellow
    } else {
        Write-Host "    [fallback] winget falhou -> tentando download direto..." -ForegroundColor DarkYellow
    }
    if (Install-Fallback $prog.id) {
        Write-Host "    [fallback] => OK" -ForegroundColor Green
        $instalados += "$($prog.nome) - Instalado (fallback direto)"
    } else {
        Write-Host "    [fallback] => ERRO (winget indisponivel e fallback falhou)" -ForegroundColor Red
        $instalados += "$($prog.nome) - ERRO (winget/fallback falharam)"
        $erros += $prog.nome
    }
}

# AnyDesk - instala em modo SILENCIOSO (sem abrir a GUI / sem clique manual)
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
        # 1) Tenta instalacao SILENCIOSA via CLI (precisa admin; --start-with-win registra o servico).
        $adDir = "${env:ProgramFiles(x86)}\AnyDesk"
        $adArgs = "--install `"$adDir`" --start-with-win --create-desktop-icon --create-start-menu-links --silent"
        Start-Process $anydeskTemp -ArgumentList $adArgs -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 5
        $anydeskServico = Get-Service -Name "AnyDesk" -ErrorAction SilentlyContinue
        # 2) FALLBACK: se o silencioso NAO registrou o servico (ja aconteceu antes), abre a GUI p/ clique manual.
        if (-not $anydeskServico) {
            Write-Host ""
            Write-Host "  ============================================" -ForegroundColor Yellow
            Write-Host "  Silencioso nao registrou o servico." -ForegroundColor Yellow
            Write-Host "  AnyDesk vai abrir: clique em 'Instalar AnyDesk'" -ForegroundColor Yellow
            Write-Host "  e depois feche a janela. O script continua." -ForegroundColor Yellow
            Write-Host "  ============================================" -ForegroundColor Yellow
            Write-Host ""
            Start-Process $anydeskTemp -Wait
            Start-Sleep -Seconds 3
            $anydeskServico = Get-Service -Name "AnyDesk" -ErrorAction SilentlyContinue
        }
        Remove-Item $anydeskTemp -Force -ErrorAction SilentlyContinue
        if ($anydeskServico) {
            Write-Host " Instalado" -ForegroundColor Green
            $instalados += "AnyDesk - Instalado"
        } else {
            Write-Host " Verificar manualmente" -ForegroundColor Yellow
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

Save-Step 5
}   # <<CKPT-CLOSE 5>>
if ($lastStep -lt 6) {   # <<CKPT-OPEN 6>>
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
    # Criar atalho em Documentos que abre a pasta config do OpenVPN
    $docsPath = [Environment]::GetFolderPath("MyDocuments")
    $configPath = "$env:ProgramFiles\OpenVPN\config"
    if (Test-Path $configPath) {
        $shell = New-Object -ComObject WScript.Shell
        $atalho = $shell.CreateShortcut("$docsPath\OpenVPN Config.lnk")
        $atalho.TargetPath = "explorer.exe"
        $atalho.Arguments = "`"$configPath`""
        $atalho.IconLocation = "$env:ProgramFiles\OpenVPN\bin\openvpn-gui.exe,0"
        $atalho.Save()
        Write-Host "  Atalho 'OpenVPN Config' criado em Documentos" -ForegroundColor Green
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

Save-Step 6
}   # <<CKPT-CLOSE 6>>
if ($lastStep -lt 7) {   # <<CKPT-OPEN 7>>
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

Save-Step 7
}   # <<CKPT-CLOSE 7>>
if ($lastStep -lt 8) {   # <<CKPT-OPEN 8>>
Write-Host "`n[8/$etapaTotal] Desativando notificacoes..." -ForegroundColor Cyan

try {
    # Desativar toasts (baloes popup) mas MANTER sons dos apps
    $regNotif = "HKCU:\Software\Microsoft\Windows\CurrentVersion\PushNotifications"
    if (-not (Test-Path $regNotif)) { New-Item -Path $regNotif -Force | Out-Null }
    Set-ItemProperty -Path $regNotif -Name "ToastEnabled" -Value 0 -Type DWord

    # Desativar Central de Notificacoes (painel lateral)
    $regAction = "HKCU:\Software\Policies\Microsoft\Windows\Explorer"
    if (-not (Test-Path $regAction)) { New-Item -Path $regAction -Force | Out-Null }
    Set-ItemProperty -Path $regAction -Name "DisableNotificationCenter" -Value 1 -Type DWord

    # Desativar toasts na tela de bloqueio, mas MANTER som
    $regLock = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings"
    if (-not (Test-Path $regLock)) { New-Item -Path $regLock -Force | Out-Null }
    Set-ItemProperty -Path $regLock -Name "NOC_GLOBAL_SETTING_ALLOW_NOTIFICATION_SOUND" -Value 1 -Type DWord
    Set-ItemProperty -Path $regLock -Name "NOC_GLOBAL_SETTING_ALLOW_TOASTS_ABOVE_LOCK" -Value 0 -Type DWord

    # Desativar sugestoes e dicas do Windows
    $regSugest = "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
    if (Test-Path $regSugest) {
        Set-ItemProperty -Path $regSugest -Name "SubscribedContent-338389Enabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $regSugest -Name "SubscribedContent-310093Enabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $regSugest -Name "SubscribedContent-338388Enabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $regSugest -Name "SoftLandingEnabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
    }

    # Desativar toasts visuais mas manter som
    $regNotifSettings = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings"
    if (-not (Test-Path $regNotifSettings)) { New-Item -Path $regNotifSettings -Force | Out-Null }
    Set-ItemProperty -Path $regNotifSettings -Name "NOC_GLOBAL_SETTING_TOASTS_ENABLED" -Value 0 -Type DWord

    # Bloquear toasts via politica
    $regPolicy = "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\PushNotifications"
    if (-not (Test-Path $regPolicy)) { New-Item -Path $regPolicy -Force | Out-Null }
    Set-ItemProperty -Path $regPolicy -Name "NoToastApplicationNotification" -Value 1 -Type DWord

    # Desativar Windows Tips/Sugestoes/Consumer Features
    $regTips = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
    if (-not (Test-Path $regTips)) { New-Item -Path $regTips -Force | Out-Null }
    Set-ItemProperty -Path $regTips -Name "DisableSoftLanding" -Value 1 -Type DWord -ErrorAction SilentlyContinue
    Set-ItemProperty -Path $regTips -Name "DisableWindowsConsumerFeatures" -Value 1 -Type DWord -ErrorAction SilentlyContinue

    Write-Host "  Toasts desativados (sons mantidos)" -ForegroundColor Green
} catch {
    Write-Host "  ERRO" -ForegroundColor Red
}

# ============================================
# [9] LIMPAR BARRA DE TAREFAS E FIXAR PROGRAMAS
# ============================================

Save-Step 8
}   # <<CKPT-CLOSE 8>>
if ($lastStep -lt 9) {   # <<CKPT-OPEN 9>>
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

    # ---- BARRA DE TAREFAS: Google Chrome + Explorador de Arquivos ----
    # Blob 'Favorites' capturado de maquina de referencia (Win11 26100). Ordem: Chrome, Explorer.
    $pinDir = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
    if (Test-Path $pinDir) { Remove-Item "$pinDir\*" -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Path $pinDir -Force | Out-Null

    # .lnk com os nomes EXATOS que o blob referencia (senao o pin nao resolve)
    $shell = New-Object -ComObject WScript.Shell
    $lnkFE = $shell.CreateShortcut("$pinDir\File Explorer.lnk")
    $lnkFE.TargetPath = "explorer.exe"
    $lnkFE.Save()
    $chromeExe = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
    if (-not (Test-Path $chromeExe)) { $chromeExe = "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe" }
    if (Test-Path $chromeExe) {
        $lnkGC = $shell.CreateShortcut("$pinDir\Google Chrome.lnk")
        $lnkGC.TargetPath = $chromeExe
        $lnkGC.Save()
    } else {
        Write-Host "  (Chrome nao encontrado - o pin do Chrome pode nao resolver)" -ForegroundColor Yellow
    }

    $taskbandPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband"
    if (-not (Test-Path $taskbandPath)) { New-Item -Path $taskbandPath -Force | Out-Null }
    # remover pins auxiliares (Copilot/TFL) e cache de resolucao antigo (Windows reconstroi)
    Remove-Item -Path "$taskbandPath\AuxilliaryPins" -Force -Recurse -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $taskbandPath -Name "FavoritesResolve" -ErrorAction SilentlyContinue

    $favBytes = [byte[]]@(
        0x00,0x50,0x01,0x00,0x00,0x3a,0x00,0x1f,0x80,0xc8,0x27,0x34,0x1f,0x10,0x5c,0x10,
        0x42,0xaa,0x03,0x2e,0xe4,0x52,0x87,0xd6,0x68,0x26,0x00,0x01,0x00,0x26,0x00,0xef,
        0xbe,0x12,0x00,0x00,0x00,0x2c,0x37,0x9c,0x99,0xf3,0x6e,0xdc,0x01,0x7c,0xa4,0xeb,
        0xc3,0xf3,0x6e,0xdc,0x01,0xe6,0xd5,0x60,0x87,0x7f,0x09,0xdd,0x01,0x14,0x00,0x56,
        0x00,0x31,0x00,0x00,0x00,0x00,0x00,0xe1,0x5c,0x8e,0x8b,0x11,0x00,0x54,0x61,0x73,
        0x6b,0x42,0x61,0x72,0x00,0x40,0x00,0x09,0x00,0x04,0x00,0xef,0xbe,0x91,0x5b,0xf2,
        0x0a,0xe1,0x5c,0x8f,0x8b,0x2e,0x00,0x00,0x00,0xc8,0x46,0x01,0x00,0x00,0x00,0x0f,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x51,
        0xbe,0x98,0x00,0x54,0x00,0x61,0x00,0x73,0x00,0x6b,0x00,0x42,0x00,0x61,0x00,0x72,
        0x00,0x00,0x00,0x16,0x00,0xbe,0x00,0x32,0x00,0xd0,0x08,0x00,0x00,0xe1,0x5c,0x03,
        0x85,0x20,0x00,0x47,0x4f,0x4f,0x47,0x4c,0x45,0x7e,0x31,0x2e,0x4c,0x4e,0x4b,0x00,
        0x00,0x54,0x00,0x09,0x00,0x04,0x00,0xef,0xbe,0xe1,0x5c,0x02,0x8c,0xe1,0x5c,0x02,
        0x8c,0x2e,0x00,0x00,0x00,0x85,0xa8,0x00,0x00,0x00,0x00,0x1f,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xb9,0x9e,0x98,0x00,0x47,
        0x00,0x6f,0x00,0x6f,0x00,0x67,0x00,0x6c,0x00,0x65,0x00,0x20,0x00,0x43,0x00,0x68,
        0x00,0x72,0x00,0x6f,0x00,0x6d,0x00,0x65,0x00,0x2e,0x00,0x6c,0x00,0x6e,0x00,0x6b,
        0x00,0x00,0x00,0x1c,0x00,0x12,0x00,0x00,0x00,0x2b,0x00,0xef,0xbe,0xe6,0xd5,0x60,
        0x87,0x7f,0x09,0xdd,0x01,0x1c,0x00,0x1a,0x00,0x00,0x00,0x1d,0x00,0xef,0xbe,0x02,
        0x00,0x43,0x00,0x68,0x00,0x72,0x00,0x6f,0x00,0x6d,0x00,0x65,0x00,0x00,0x00,0x1c,
        0x00,0x22,0x00,0x00,0x00,0x1e,0x00,0xef,0xbe,0x02,0x00,0x55,0x00,0x73,0x00,0x65,
        0x00,0x72,0x00,0x50,0x00,0x69,0x00,0x6e,0x00,0x6e,0x00,0x65,0x00,0x64,0x00,0x00,
        0x00,0x1c,0x00,0x00,0x00,0x00,0xa0,0x01,0x00,0x00,0x3a,0x00,0x1f,0x80,0xc8,0x27,
        0x34,0x1f,0x10,0x5c,0x10,0x42,0xaa,0x03,0x2e,0xe4,0x52,0x87,0xd6,0x68,0x26,0x00,
        0x01,0x00,0x26,0x00,0xef,0xbe,0x12,0x00,0x00,0x00,0x2c,0x37,0x9c,0x99,0xf3,0x6e,
        0xdc,0x01,0x7c,0xa4,0xeb,0xc3,0xf3,0x6e,0xdc,0x01,0x31,0xbf,0x58,0x8b,0x7f,0x09,
        0xdd,0x01,0x14,0x00,0x56,0x00,0x31,0x00,0x00,0x00,0x00,0x00,0xe1,0x5c,0x05,0x8c,
        0x11,0x00,0x54,0x61,0x73,0x6b,0x42,0x61,0x72,0x00,0x40,0x00,0x09,0x00,0x04,0x00,
        0xef,0xbe,0x91,0x5b,0xf2,0x0a,0xe1,0x5c,0x05,0x8c,0x2e,0x00,0x00,0x00,0xc8,0x46,
        0x01,0x00,0x00,0x00,0x0f,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
        0x00,0x00,0x00,0x00,0x81,0x3b,0x56,0x00,0x54,0x00,0x61,0x00,0x73,0x00,0x6b,0x00,
        0x42,0x00,0x61,0x00,0x72,0x00,0x00,0x00,0x16,0x00,0x0e,0x01,0x32,0x00,0x97,0x01,
        0x00,0x00,0x81,0x58,0xc4,0x3a,0x20,0x00,0x46,0x49,0x4c,0x45,0x45,0x58,0x7e,0x31,
        0x2e,0x4c,0x4e,0x4b,0x00,0x00,0x7c,0x00,0x09,0x00,0x04,0x00,0xef,0xbe,0xe1,0x5c,
        0x05,0x8c,0xe1,0x5c,0x05,0x8c,0x2e,0x00,0x00,0x00,0xa9,0x81,0x02,0x00,0x00,0x00,
        0x07,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x52,0x00,0x00,0x00,0x00,0x00,
        0xdb,0xdc,0x91,0x00,0x46,0x00,0x69,0x00,0x6c,0x00,0x65,0x00,0x20,0x00,0x45,0x00,
        0x78,0x00,0x70,0x00,0x6c,0x00,0x6f,0x00,0x72,0x00,0x65,0x00,0x72,0x00,0x2e,0x00,
        0x6c,0x00,0x6e,0x00,0x6b,0x00,0x00,0x00,0x40,0x00,0x73,0x00,0x68,0x00,0x65,0x00,
        0x6c,0x00,0x6c,0x00,0x33,0x00,0x32,0x00,0x2e,0x00,0x64,0x00,0x6c,0x00,0x6c,0x00,
        0x2c,0x00,0x2d,0x00,0x32,0x00,0x32,0x00,0x30,0x00,0x36,0x00,0x37,0x00,0x00,0x00,
        0x1c,0x00,0x12,0x00,0x00,0x00,0x2b,0x00,0xef,0xbe,0x0a,0xa0,0x59,0x8b,0x7f,0x09,
        0xdd,0x01,0x1c,0x00,0x42,0x00,0x00,0x00,0x1d,0x00,0xef,0xbe,0x02,0x00,0x4d,0x00,
        0x69,0x00,0x63,0x00,0x72,0x00,0x6f,0x00,0x73,0x00,0x6f,0x00,0x66,0x00,0x74,0x00,
        0x2e,0x00,0x57,0x00,0x69,0x00,0x6e,0x00,0x64,0x00,0x6f,0x00,0x77,0x00,0x73,0x00,
        0x2e,0x00,0x45,0x00,0x78,0x00,0x70,0x00,0x6c,0x00,0x6f,0x00,0x72,0x00,0x65,0x00,
        0x72,0x00,0x00,0x00,0x1c,0x00,0x22,0x00,0x00,0x00,0x1e,0x00,0xef,0xbe,0x02,0x00,
        0x55,0x00,0x73,0x00,0x65,0x00,0x72,0x00,0x50,0x00,0x69,0x00,0x6e,0x00,0x6e,0x00,
        0x65,0x00,0x64,0x00,0x00,0x00,0x1c,0x00,0x00,0x00,0xff
    )
    Set-ItemProperty -Path $taskbandPath -Name "Favorites" -Value $favBytes -Type Binary
    Set-ItemProperty -Path $taskbandPath -Name "FavoritesVersion" -Value 3 -Type DWord -ErrorAction SilentlyContinue

    # reiniciar o explorer p/ aplicar a barra
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    if (-not (Get-Process -Name explorer -ErrorAction SilentlyContinue)) { Start-Process explorer.exe }

    Write-Host "  Barra padronizada (Google Chrome + Explorador de Arquivos)" -ForegroundColor Green
} catch {
    Write-Host "  ERRO" -ForegroundColor Red
}

# ============================================
# [10] WALLPAPER IG NETWORKS
# ============================================

Save-Step 9
}   # <<CKPT-CLOSE 9>>
if ($lastStep -lt 10) {   # <<CKPT-OPEN 10>>
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

Save-Step 10
}   # <<CKPT-CLOSE 10>>
if ($lastStep -lt 11) {   # <<CKPT-OPEN 11>>
Write-Host "`n[11/$etapaTotal] Configurando energia..." -ForegroundColor Cyan

# AC (carregador): nada apaga / nao dorme
powercfg /change monitor-timeout-ac 0 2>&1 | Out-Null
powercfg /change standby-timeout-ac 0 2>&1 | Out-Null
powercfg /change hibernate-timeout-ac 0 2>&1 | Out-Null
Write-Host "  AC (carregador): nunca apaga tela, nunca dorme" -ForegroundColor Green

# Bateria: tela apaga em 30 min, NAO dorme
powercfg /change monitor-timeout-dc 30 2>&1 | Out-Null
powercfg /change standby-timeout-dc 0 2>&1 | Out-Null
powercfg /change hibernate-timeout-dc 0 2>&1 | Out-Null
Write-Host "  Bateria: tela apaga em 30 min, sem sleep" -ForegroundColor Green

# Botoes: 0=nada, 1=sleep, 2=hibernate, 3=shutdown, 4=desliga display
# Power button -> shutdown (real, sem Fast Startup)
powercfg /setacvalueindex SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 7648efa3-dd9c-4e3e-b566-50f929386280 3 2>&1 | Out-Null
powercfg /setdcvalueindex SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 7648efa3-dd9c-4e3e-b566-50f929386280 3 2>&1 | Out-Null
# Lid close -> sleep (volta quando abrir a tampa)
powercfg /setacvalueindex SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 1 2>&1 | Out-Null
powercfg /setdcvalueindex SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 1 2>&1 | Out-Null
Write-Host "  Botao Power = desligar | Fechar tampa = sleep (volta ao abrir)" -ForegroundColor Green

# Desabilita hibernacao por completo (tambem desabilita Fast Startup)
powercfg /hibernate off 2>&1 | Out-Null
REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power" /V HiberbootEnabled /T REG_DWORD /D 0 /F 2>&1 | Out-Null
Write-Host "  Hibernacao desabilitada (shutdown = shutdown real, sem Fast Startup)" -ForegroundColor Green

# Desabilita Modern Standby (S0ix) - causa do "LED piscando" apos shutdown em notebooks novos.
# Forca o sistema a usar S3 (sleep) e S5 (shutdown real) tradicionais.
REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Power" /V PlatformAoAcOverride /T REG_DWORD /D 0 /F 2>&1 | Out-Null
Write-Host "  Modern Standby desabilitado (S5 forcado no shutdown)" -ForegroundColor Green

# Desabilita TODOS os devices que podem acordar o PC (Wake on LAN, USB, etc).
# Esses devices mantem rail de energia ligado mesmo apos shutdown, causando LED ativo.
try {
    $wakeDevices = powercfg /devicequery wake_armed 2>&1 | Where-Object { $_ -and $_ -notmatch 'NONE|------' }
    $disabled = 0
    foreach ($dev in $wakeDevices) {
        $d = "$dev".Trim()
        if ($d) {
            powercfg /devicedisablewake "$d" 2>&1 | Out-Null
            $disabled++
        }
    }
    if ($disabled -gt 0) { Write-Host "  $disabled dispositivo(s) impedido(s) de acordar o PC" -ForegroundColor Green }
} catch {}

# Aplica
powercfg /setactive SCHEME_CURRENT 2>&1 | Out-Null


Save-Step 11
}   # <<CKPT-CLOSE 11>>
if ($lastStep -lt 12) {   # <<CKPT-OPEN 12>>
Write-Host "`n[12/$etapaTotal] Removendo programas do inicio automatico..." -ForegroundColor Cyan

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

Save-Step 12
}   # <<CKPT-CLOSE 12>>
if ($lastStep -lt 13) {   # <<CKPT-OPEN 13>>
Write-Host "`n[13/$etapaTotal] Verificando conta do usuario..." -ForegroundColor Cyan

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

Save-Step 13
}   # <<CKPT-CLOSE 13>>
# Fecha o transcript do log antes do pause
if ($logAtivo) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor DarkGray
    Write-Host "  LOG SALVO EM:" -ForegroundColor Green
    Write-Host "  $logPath" -ForegroundColor Green
    Write-Host "  (Se algo der errado, mande esse arquivo para diagnostico)" -ForegroundColor DarkGray
    Write-Host "========================================" -ForegroundColor DarkGray
    try { Stop-Transcript | Out-Null } catch {}
}

# checkpoint concluido -> limpar p/ proxima maquina comecar do zero
try { Remove-Item $ckptFile -Force -ErrorAction SilentlyContinue } catch {}
pause
