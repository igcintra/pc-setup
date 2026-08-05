# ============================================================
#  pos-uti.ps1 — fechamento de atendimento depois da UTI
#
#  Resolve o caso "reiniciei mas o PC nao reiniciou de verdade":
#  com Fast Startup ligado, DESLIGAR+LIGAR faz desligamento hibrido
#  (o kernel volta do disco) -> o uptime NAO zera e as pendencias de
#  DISM/SFC/updates NAO sao aplicadas. Só "Reiniciar" aplica.
#
#  Faz, em ordem:
#    1. Mostra uptime REAL e ultimo boot
#    2. Diz QUAIS flags de reinicio pendente existem
#    3. Desliga o Fast Startup (garantia)
#    4. Ajusta espaco das Copias de Sombra (erro Volsnap)
#    5. Tenta destravar o servico do Google Update (gupdate)
#    6. Lista adaptadores de rede com problema (Wi-Fi Direct)
#    7. Poe o PC no canal ESTAVEL de updates (sai da build antecipada)
#    8. Oferece REINICIAR de verdade
#
#  Uso (PowerShell; se nao for admin ele se auto-eleva):
#    irm https://raw.githubusercontent.com/igcintra/pc-setup/master/pos-uti.ps1 | iex
# ============================================================

$URL = 'https://raw.githubusercontent.com/igcintra/pc-setup/master/pos-uti.ps1'
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host 'Elevando para Administrador...' -ForegroundColor Yellow
    Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-Command',"irm $URL | iex"
    return
}

Write-Host "`n=== POS-UTI: fechamento do atendimento ===" -ForegroundColor Cyan

# ---------- 1. uptime real ----------
$os = Get-CimInstance Win32_OperatingSystem
$boot = $os.LastBootUpTime
$up = (Get-Date) - $boot
Write-Host ("`n[1] Ultimo boot: {0:dd/MM/yyyy HH:mm}  (ligado ha {1} dias e {2}h)" -f $boot, $up.Days, $up.Hours)
if ($up.TotalDays -ge 1) {
    Write-Host "    ATENCAO: mais de 1 dia sem reiniciar de verdade." -ForegroundColor Yellow
    Write-Host "    Se voce 'desligou e ligou' nesse periodo, o Fast Startup impediu o reinicio real." -ForegroundColor Yellow
}

# ---------- 2. o que esta pendente ----------
Write-Host "`n[2] Reinicio pendente:" -ForegroundColor Cyan
$pend = @()
if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $pend += 'reparos do Windows (CBS)' }
if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $pend += 'Windows Update' }
if (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue) { $pend += 'arquivos aguardando troca no boot (PendingFileRename)' }
$cn = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName' -Name ComputerName -ErrorAction SilentlyContinue
$cn2 = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName' -Name ComputerName -ErrorAction SilentlyContinue
if ($cn -and $cn2 -and $cn.ComputerName -ne $cn2.ComputerName) { $pend += 'renomeacao do computador' }
if ($pend.Count) { $pend | ForEach-Object { Write-Host "    - $_" -ForegroundColor Yellow } }
else { Write-Host "    nada pendente" -ForegroundColor Green }

# ---------- 3. Fast Startup OFF ----------
Write-Host "`n[3] Fast Startup (inicializacao rapida):" -ForegroundColor Cyan
$rp = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power'
$antes = (Get-ItemProperty $rp -Name HiberbootEnabled -ErrorAction SilentlyContinue).HiberbootEnabled
Set-ItemProperty -Path $rp -Name HiberbootEnabled -Value 0 -Type DWord
Write-Host ("    estava " + $(if ($antes -eq 1) {'LIGADO -> desligado agora'} else {'ja desligado'})) -ForegroundColor Green
Write-Host "    (com ele ligado, 'desligar e ligar' NAO aplica reparos — so 'Reiniciar')" -ForegroundColor Yellow

# ---------- 4. Copias de Sombra (Volsnap) ----------
Write-Host "`n[4] Copias de Sombra / Restauracao do Sistema:" -ForegroundColor Cyan
try {
    $out = vssadmin resize shadowstorage /for=C: /on=C: /maxsize=10% 2>&1
    if ($out -match 'com ˆxito|successfully|com exito|com êxito') { Write-Host "    espaco ajustado para 10% do disco" -ForegroundColor Green }
    else { Write-Host "    (nao alterou — pode ser que a Protecao do Sistema esteja desativada)" -ForegroundColor Yellow }
} catch { Write-Host "    nao consegui ajustar" -ForegroundColor Yellow }

# ---------- 5. Google Update ----------
Write-Host "`n[5] Servico do Google Update (gupdate):" -ForegroundColor Cyan
$g = Get-Service gupdate -ErrorAction SilentlyContinue
if ($g) {
    try {
        Set-Service gupdate -StartupType Manual -ErrorAction Stop
        Start-Service gupdate -ErrorAction Stop
        Write-Host "    ok, subiu normalmente" -ForegroundColor Green
    } catch {
        Write-Host "    NAO sobe — instalacao do Chrome provavelmente danificada." -ForegroundColor Yellow
        Write-Host "    Acao: reinstalar o Google Chrome (baixar do site oficial e instalar por cima)." -ForegroundColor Yellow
    }
} else { Write-Host "    servico nao existe nesta maquina (ok se nao usa Chrome)" }

# ---------- 6. adaptadores de rede com erro ----------
Write-Host "`n[6] Adaptadores de rede com problema:" -ForegroundColor Cyan
$bad = Get-PnpDevice -Class Net -ErrorAction SilentlyContinue | Where-Object { $_.Status -ne 'OK' }
if ($bad) {
    $bad | ForEach-Object { Write-Host ("    [" + $_.Status + "] " + $_.FriendlyName) -ForegroundColor Yellow }
    Write-Host "    Acao: atualizar o driver de Wi-Fi (Lenovo Vantage / Dell Command / site do fabricante)." -ForegroundColor Yellow
} else { Write-Host "    todos OK" -ForegroundColor Green }

# ---------- 7. canal estavel de updates ----------
Write-Host "`n[7] Canal de atualizacoes:" -ForegroundColor Cyan
$wu = 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings'
if (-not (Test-Path $wu)) { New-Item -Path $wu -Force | Out-Null }
$ant = (Get-ItemProperty $wu -Name IsContinuousInnovationOptedIn -ErrorAction SilentlyContinue).IsContinuousInnovationOptedIn
Set-ItemProperty -Path $wu -Name IsContinuousInnovationOptedIn -Value 0 -Type DWord
Write-Host ("    updates antecipados: " + $(if ($ant -eq 1) {'ESTAVAM LIGADOS -> desligados agora'} else {'ja estavam desligados'})) -ForegroundColor Green
$b = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
Write-Host ("    build atual: " + $b.CurrentBuildNumber + "." + $b.UBR + " (a atual nao volta atras; para de antecipar as proximas)")

# ---------- 8. reiniciar de verdade ----------
Write-Host "`n=== TUDO APLICADO ===" -ForegroundColor Green
Write-Host "Falta o passo mais importante: REINICIAR DE VERDADE." -ForegroundColor White
Write-Host "Depois do reinicio, rode a UTI novamente e NAO use o PC durante ela." -ForegroundColor White
$r = Read-Host "`nReiniciar agora? (S/N)"
if ($r -match '^[SsYy]') {
    Write-Host "Reiniciando em 30 segundos... (para cancelar: shutdown /a)" -ForegroundColor Yellow
    shutdown /r /f /t 30 /c "TI: reinicio para aplicar reparos do Windows (DISM/SFC/updates)"
} else {
    Write-Host "OK — mas reinicie assim que puder, usando INICIAR > Reiniciar (nao 'Desligar')." -ForegroundColor Yellow
}
