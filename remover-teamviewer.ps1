# ============================================================================
# remover-teamviewer.ps1 — IG Networks / TI
# ----------------------------------------------------------------------------
# Remove o TeamViewer da maquina (instalado ou modulo QuickSupport no perfil do
# usuario), junto com tarefa agendada, chaves de inicializacao e pastas.
#
# GUARDA A PROVA ANTES DE APAGAR: gera um TXT com inventario (caminhos, hash,
# quem assinou, datas, ID do modulo customizado) e um ZIP da pasta, os dois no
# Desktop. Isso importa porque o modulo QS e' CUSTOMIZADO por um terceiro -- o
# ID identifica de quem e', e depois de apagar nao da' mais para saber.
#
# NAO TOCA NO ANYDESK (ferramenta oficial da casa).
#
# Nao exige administrador. Se a tarefa agendada tiver sido criada por admin, a
# remocao dela falha e o script avisa -- o resto e' feito igual.
#
# Uso (pede confirmacao na tela, e' digitar REMOVER):
#   irm https://raw.githubusercontent.com/igcintra/pc-setup/master/remover-teamviewer.ps1 | iex
#
# Uso sem perguntar (TI, ja decidido):
#   & ([scriptblock]::Create((irm https://raw.githubusercontent.com/igcintra/pc-setup/master/remover-teamviewer.ps1))) -Doit
# ============================================================================
param([switch]$Doit)

$ErrorActionPreference = 'SilentlyContinue'
$RX = 'teamviewer|tv_w32|tv_x64'      # nunca inclui anydesk

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  REMOVER TEAMVIEWER - IG Networks" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""

$det   = New-Object System.Collections.ArrayList
$achou = @()

# ArrayList de proposito: .Add() muta o MESMO objeto, entao funciona igual rodando
# via 'irm | iex' ou via scriptblock (com array + '+=' o escopo pode furar).
function Add-Det($linha) { [void]$det.Add([string]$linha) }

# ---------------- 1) INVENTARIO (o que existe agora) ----------------
Write-Host "[1/4] Procurando o TeamViewer..." -ForegroundColor White

$pastas = @()
foreach ($base in @($env:LOCALAPPDATA, $env:APPDATA, $env:ProgramData, $env:ProgramFiles, ${env:ProgramFiles(x86)})) {
    if ($base) {
        $p = Join-Path $base 'TeamViewer'
        if (Test-Path -LiteralPath $p) { $pastas += $p }
    }
}

$procs = @(Get-Process | Where-Object { $_.Name -match $RX })
$svcs  = @(Get-CimInstance Win32_Service | Where-Object { $_.Name -match $RX -or $_.PathName -match 'TeamViewer' })
$tasks = @(Get-ScheduledTask | Where-Object { $_.TaskName -match $RX -or (($_.Actions | Where-Object { $_.Execute }).Execute -join ' ') -match 'TeamViewer' })

$raizes = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*')
$instalados = @(Get-ItemProperty $raizes | Where-Object { $_.DisplayName -match 'TeamViewer' })

$runKeys = @()
foreach ($k in @('HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
                 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
                 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run')) {
    if (Test-Path $k) {
        (Get-Item $k).Property | ForEach-Object {
            $v = (Get-ItemProperty $k -Name $_).$_
            if ("$_ $v" -match 'TeamViewer') { $runKeys += [pscustomobject]@{ Chave=$k; Nome=$_; Valor=$v } }
        }
    }
}
$startup = @(Get-ChildItem (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup') -File |
             Where-Object { $_.Name -match 'teamviewer' })
$regTV = @('HKCU:\SOFTWARE\TeamViewer','HKLM:\SOFTWARE\TeamViewer','HKLM:\SOFTWARE\WOW6432Node\TeamViewer') |
         Where-Object { Test-Path $_ }

Add-Det "----- INVENTARIO -----"
foreach ($p in $pastas) {
    $achou += "pasta"
    $it = Get-ChildItem -LiteralPath $p -Recurse -File
    $mais = $it | Sort-Object CreationTime | Select-Object -First 1
    Add-Det "[pasta] $p"
    Add-Det "        $($it.Count) arquivos, criada/1o arquivo em $($mais.CreationTime)"
    # ID do modulo customizado -- e' o que identifica DE QUEM e' o QuickSupport
    Get-ChildItem -LiteralPath (Join-Path $p 'CustomConfigs') -Directory | ForEach-Object {
        Add-Det "        MODULO CUSTOMIZADO id=$($_.Name)  (criado $($_.CreationTime))"
    }
    Get-ChildItem -LiteralPath $p -Recurse -Filter 'TeamViewer*.exe' | Select-Object -First 5 | ForEach-Object {
        $sig = Get-AuthenticodeSignature -LiteralPath $_.FullName
        $cn = ''
        if ($sig.SignerCertificate) { $cn = (($sig.SignerCertificate.Subject -split ',' | Where-Object { $_ -match 'CN=' } | Select-Object -First 1) -replace '.*CN=','').Trim() }
        Add-Det "        exe: $($_.FullName)"
        Add-Det "             SHA256=$((Get-FileHash $_.FullName -Algorithm SHA256).Hash)"
        Add-Det "             assinatura=$($sig.Status) assinante=$cn  data=$($_.LastWriteTime)"
    }
    # log de conexoes: quem entrou nessa maquina (fica no ZIP)
    Get-ChildItem -LiteralPath $p -Recurse -Include 'Connections*.txt','TeamViewer*_Logfile*.log' | ForEach-Object {
        Add-Det "        LOG DE CONEXOES: $($_.Name) ($([int]($_.Length/1024)) KB) -> vai no ZIP"
    }
}
foreach ($t in $tasks)      { $achou += "tarefa";   Add-Det "[tarefa] $($t.TaskName) [$($t.State)] -> $((($t.Actions | Where-Object { $_.Execute }).Execute -join ' '))" }
foreach ($s in $svcs)       { $achou += "servico";  Add-Det "[servico] $($s.Name) [$($s.State)] -> $($s.PathName)" }
foreach ($i in $instalados) { $achou += "programa"; Add-Det "[programa] $($i.DisplayName) $($i.DisplayVersion) -> uninstall: $($i.UninstallString)" }
foreach ($r in $runKeys)    { $achou += "run";      Add-Det "[inicializacao] $($r.Chave)\$($r.Nome) = $($r.Valor)" }
foreach ($s in $startup)    { $achou += "startup";  Add-Det "[startup] $($s.FullName)" }
foreach ($r in $regTV)      { $achou += "registro"; Add-Det "[registro] $r" }
foreach ($p in $procs)      { $achou += "processo"; Add-Det "[EM EXECUCAO] $($p.Name) pid=$($p.Id) -> $($p.Path)" }

$achou = $achou | Select-Object -Unique

if (-not $achou.Count) {
    Write-Host ""
    Write-Host "Nao encontrei TeamViewer nesta maquina. Nada a fazer." -ForegroundColor Green
    Write-Host ""
    return
}

Write-Host "      Encontrado: $($achou -join ', ')" -ForegroundColor Yellow

# ---------------- 2) GUARDAR A PROVA ----------------
Write-Host "[2/4] Guardando a prova no Desktop (antes de apagar)..." -ForegroundColor White

$desktop = [Environment]::GetFolderPath("Desktop")
if (-not $desktop -or -not (Test-Path $desktop)) { $desktop = "$env:USERPROFILE\Desktop" }
$stamp   = Get-Date -Format 'yyyy-MM-dd'
$limpaU  = ($env:USERNAME     -replace '[^A-Za-z0-9._-]','')
$limpaH  = ($env:COMPUTERNAME -replace '[^A-Za-z0-9._-]','')
$txt     = Join-Path $desktop "teamviewer-remocao-$limpaU-$limpaH-$stamp.txt"
$zip     = Join-Path $desktop "teamviewer-evidencia-$limpaU-$limpaH-$stamp.zip"

$cab = @(
    "REMOCAO DE TEAMVIEWER - IG Networks",
    "Data......: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')",
    "Computador: $env:COMPUTERNAME",
    "Usuario...: $env:USERNAME",
    "Windows...: $((Get-CimInstance Win32_OperatingSystem).Caption) build $((Get-CimInstance Win32_OperatingSystem).BuildNumber)",
    "Encontrado: $($achou -join ', ')",
    ""
)
($cab + $det) | Set-Content -Path $txt -Encoding UTF8

$paraZipar = @($pastas | Where-Object { $_ -match 'AppData|ProgramData' })
if ($paraZipar.Count) {
    try {
        Compress-Archive -Path $paraZipar -DestinationPath $zip -Force -ErrorAction Stop
        Write-Host "      ZIP salvo ($([int]((Get-Item $zip).Length/1MB)) MB)" -ForegroundColor Gray
    } catch {
        # TeamViewer aberto mantem o log preso e o ZIP falha
        Add-Det "[aviso] o ZIP da prova FALHOU: $($_.Exception.Message)"
        Write-Host "      ATENCAO: nao consegui zipar a pasta (TeamViewer aberto?)." -ForegroundColor Yellow
        Write-Host "      FECHE o TeamViewer e rode de novo, para a prova nao se perder." -ForegroundColor Yellow
        if ($procs.Count) {
            Write-Host "      (esta em execucao agora: $(($procs | ForEach-Object { $_.Name }) -join ', '))" -ForegroundColor Yellow
            Write-Host ""
            $c = Read-Host "Continuar SEM a copia da pasta? (digite SIM para continuar)"
            if ($c -ne 'SIM') {
                ($cab + $det) | Set-Content -Path $txt -Encoding UTF8
                Write-Host "Parado. Nada foi removido. Feche o TeamViewer e rode de novo." -ForegroundColor Yellow
                return
            }
        }
    }
} else {
    Write-Host "      (instalacao em Program Files: nao zipa, o desinstalador cuida)" -ForegroundColor Gray
}
Write-Host "      Relatorio: $(Split-Path $txt -Leaf)" -ForegroundColor Gray

# ---------------- 3) CONFIRMAR ----------------
if (-not $Doit) {
    Write-Host ""
    Write-Host "O que sera removido:" -ForegroundColor Yellow
    foreach ($l in $det | Where-Object { $_ -match '^\[' }) { Write-Host "   $l" -ForegroundColor Gray }
    Write-Host ""
    Write-Host "O AnyDesk NAO sera tocado." -ForegroundColor Green
    Write-Host ""
    $resp = Read-Host "Para remover, digite REMOVER e aperte Enter (qualquer outra coisa cancela)"
    if ($resp -ne 'REMOVER') {
        Write-Host ""
        Write-Host "Cancelado. Nada foi removido. A prova ficou salva no Desktop." -ForegroundColor Yellow
        Write-Host ""
        return
    }
}

# ---------------- 4) REMOVER ----------------
Write-Host "[3/4] Removendo..." -ForegroundColor White
$feito = @(); $falhou = @()

foreach ($p in $procs) {
    Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
    if (Get-Process -Id $p.Id -ErrorAction SilentlyContinue) { $falhou += "processo $($p.Name)" } else { $feito += "processo $($p.Name) encerrado" }
}
Start-Sleep -Seconds 2

foreach ($t in $tasks) {
    Unregister-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -Confirm:$false -ErrorAction SilentlyContinue
    if (Get-ScheduledTask -TaskName $t.TaskName -ErrorAction SilentlyContinue) {
        $falhou += "tarefa $($t.TaskName) (provavelmente criada por administrador)"
    } else { $feito += "tarefa $($t.TaskName) removida" }
}

foreach ($s in $svcs) {
    Stop-Service -Name $s.Name -Force -ErrorAction SilentlyContinue
    & sc.exe delete $s.Name | Out-Null
    if (Get-Service -Name $s.Name -ErrorAction SilentlyContinue) { $falhou += "servico $($s.Name) (precisa admin)" } else { $feito += "servico $($s.Name) removido" }
}

foreach ($i in $instalados) {
    $u = $i.UninstallString
    if ($u) {
        $feito += "desinstalador chamado: $($i.DisplayName)"
        if ($u -match '^"?([^"]+\.exe)"?\s*(.*)$') {
            Start-Process -FilePath $matches[1] -ArgumentList ("$($matches[2]) /S").Trim() -Wait -ErrorAction SilentlyContinue
        }
    }
}

foreach ($r in $runKeys) { Remove-ItemProperty -Path $r.Chave -Name $r.Nome -Force -ErrorAction SilentlyContinue; $feito += "inicializacao $($r.Nome) removida" }
foreach ($s in $startup) { Remove-Item -LiteralPath $s.FullName -Force -ErrorAction SilentlyContinue; $feito += "atalho de startup removido" }
foreach ($r in $regTV)   { Remove-Item -Path $r -Recurse -Force -ErrorAction SilentlyContinue; if (Test-Path $r) { $falhou += "registro $r" } else { $feito += "registro $r removido" } }

foreach ($p in $pastas) {
    Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $p) { $falhou += "pasta $p (arquivo em uso?)" } else { $feito += "pasta $p removida" }
}

# ---------------- confere de novo ----------------
Write-Host "[4/4] Conferindo..." -ForegroundColor White
$sobrou = @()
foreach ($base in @($env:LOCALAPPDATA, $env:APPDATA, $env:ProgramData, $env:ProgramFiles, ${env:ProgramFiles(x86)})) {
    if ($base -and (Test-Path -LiteralPath (Join-Path $base 'TeamViewer'))) { $sobrou += (Join-Path $base 'TeamViewer') }
}
if (Get-ScheduledTask | Where-Object { $_.TaskName -match $RX })      { $sobrou += "tarefa agendada" }
if (Get-Process | Where-Object { $_.Name -match $RX })                { $sobrou += "processo em execucao" }

Add-Det ""
Add-Det "----- REMOCAO -----"
$feito  | ForEach-Object { Add-Det "[ok] $_" }
$falhou | ForEach-Object { Add-Det "[FALHOU] $_" }
$sobrou | ForEach-Object { Add-Det "[SOBROU] $_" }
($cab + $det) | Set-Content -Path $txt -Encoding UTF8

Write-Host ""
if ($sobrou.Count -or $falhou.Count) {
    Write-Host "=========================================" -ForegroundColor Yellow
    Write-Host "  REMOVIDO EM PARTE - AVISE A TI" -ForegroundColor Yellow
    Write-Host "=========================================" -ForegroundColor Yellow
    $falhou | ForEach-Object { Write-Host "  falhou: $_" -ForegroundColor Gray }
    $sobrou | ForEach-Object { Write-Host "  sobrou: $_" -ForegroundColor Gray }
    Write-Host ""
    Write-Host "Sobrou coisa que precisa de administrador. Reinicie o PC e rode de novo," -ForegroundColor Gray
    Write-Host "ou chame a TI." -ForegroundColor Gray
} else {
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host "  TEAMVIEWER REMOVIDO" -ForegroundColor Green
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "O AnyDesk continua funcionando normalmente." -ForegroundColor Gray
}
Write-Host ""
Write-Host "Mande estes 2 arquivos do seu Desktop para a TI:" -ForegroundColor Yellow
Write-Host "  $(Split-Path $txt -Leaf)" -ForegroundColor White
if (Test-Path $zip) { Write-Host "  $(Split-Path $zip -Leaf)" -ForegroundColor White }
Write-Host ""
