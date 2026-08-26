# ============================================================
#  corrigir-hibernacao.ps1  —  devolve o HIBERNAR ao notebook
#
#  SINTOMA: em "Escolher o que fazer ao fechar a tampa" so aparecem
#  "Nao fazer nada" e "Desligar" — sem Suspender e sem Hibernar.
#  Resultado: a tampa fecha e o PC continua LIGADO -> bateria acaba
#  durante a noite (chega em 10% no dia seguinte).
#
#  CAUSA (eh o nosso proprio setup.ps1, etapa [11] "Configurando energia"):
#    a) `powercfg /hibernate off`      -> tira o HIBERNAR da lista
#    b) `PlatformAoAcOverride = 0`     -> desliga o Modern Standby (S0ix).
#       Em notebook novo que so tem S0ix no firmware (nao tem S3), isso
#       tira TAMBEM o SUSPENDER -> sobra so "nada" e "desligar".
#    (a) foi para o shutdown ser real e (b) para o LED nao ficar piscando
#    depois de desligar. Os dois efeitos colaterais moram nesta tela.
#
#  O QUE ESTE SCRIPT FAZ: religa a hibernacao (S4, que NAO depende de
#  S3/S0ix), mantem o Fast Startup desligado, e poe "fechar a tampa =
#  hibernar". Assim a tampa fechada gasta ZERO bateria e o trabalho volta
#  do jeito que estava. O Modern Standby continua desligado (LED ok).
#
#  Caso de origem: Pablo Bencardino (26/08/2026)
#
#  Uso (PowerShell, auto-eleva):
#    irm https://raw.githubusercontent.com/igcintra/pc-setup/master/corrigir-hibernacao.ps1 | iex
#  ou local:  powershell -ExecutionPolicy Bypass -File .\corrigir-hibernacao.ps1
# ============================================================

$URL = 'https://raw.githubusercontent.com/igcintra/pc-setup/master/corrigir-hibernacao.ps1'

# ---- auto-elevacao (re-baixa a si mesmo elevado, sem arquivo) ----
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host 'Elevando para Administrador...' -ForegroundColor Yellow
    Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-Command',"irm $URL | iex"
    return
}

$Host.UI.RawUI.WindowTitle = 'Corrigindo hibernacao - nao feche esta janela'
$logDir = Join-Path ([Environment]::GetFolderPath('Desktop')) 'UTI-backup'
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
Start-Transcript -Path (Join-Path $logDir 'corrigir-hibernacao-log.txt') -Append | Out-Null

Write-Host ""
Write-Host "===== corrigir-hibernacao - $(Get-Date) =====" -ForegroundColor Cyan
Write-Host "PC: $env:COMPUTERNAME   Usuario: $env:USERNAME" -ForegroundColor DarkGray

# GUIDs do subgrupo "Botoes de energia e tampa"
$SUB_BUTTONS = '4f971e89-eebd-4455-a8de-9e59040e7347'
$GUID_LID    = '5ca83367-6e45-459f-a27b-476b1d01c936'   # acao ao fechar a tampa
$GUID_POWER  = '7648efa3-dd9c-4e3e-b566-50f929386280'   # botao de energia
# valores: 0=nada  1=suspender  2=hibernar  3=desligar  4=desliga display

# ------------------------------------------------------------
# 1) ANTES: quais estados de energia a maquina oferece hoje
# ------------------------------------------------------------
Write-Host "`n[1/6] Estados disponiveis ANTES:" -ForegroundColor Cyan
$antes = (powercfg /a) -join "`n"
Write-Host $antes -ForegroundColor DarkGray

$ram = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 0)
$hiberGB = [math]::Round($ram * 0.4, 1)
Write-Host "  [i] o hiberfil.sys vai ocupar ~$hiberGB GB no C: (40% de $ram GB de RAM)" -ForegroundColor Yellow

# ------------------------------------------------------------
# 2) Religar a hibernacao (tipo FULL — 'reduced' nao hiberna, so Fast Startup)
# ------------------------------------------------------------
Write-Host "`n[2/6] Religando a hibernacao..." -ForegroundColor Cyan
powercfg /hibernate on 2>&1 | Out-Null
powercfg /hibernate /type full 2>&1 | Out-Null

$hiberFile = Join-Path $env:SystemDrive 'hiberfil.sys'

# A prova imediata eh o registro; o hiberfil.sys de vários GB pode levar segundos
# para aparecer (no 1o caso real o script todo rodou em 1s e reportou falso alarme).
function Test-HibernacaoLigada {
    try {
        $v = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Power' `
              -Name HibernateEnabled -ErrorAction Stop).HibernateEnabled
        return [int]$v -eq 1
    } catch { return $false }
}

$regOk = Test-HibernacaoLigada
$temHiberfil = $false
for ($i = 0; $i -lt 10; $i++) {                 # ate 5s esperando o arquivo aparecer
    if (Test-Path $hiberFile) { $temHiberfil = $true; break }
    Start-Sleep -Milliseconds 500
}

if ($regOk -or $temHiberfil) {
    $tam = ''
    if ($temHiberfil) {
        try { $tam = ' (' + [math]::Round((Get-Item $hiberFile -Force).Length / 1GB, 1) + ' GB)' } catch {}
    }
    Write-Host "  [ok] hibernacao LIGADA$tam" -ForegroundColor Green
    if (-not $temHiberfil) {
        Write-Host "  [i] o hiberfil.sys ainda nao aparece no disco - normal, o Windows o cria em segundo plano" -ForegroundColor DarkGray
    }
} else {
    Write-Host "  [!] a hibernacao NAO ligou - ver o motivo na saida do passo 6" -ForegroundColor Red
}

# ------------------------------------------------------------
# 3) Fast Startup CONTINUA desligado (o /hibernate on o religa sozinho)
#    Sem isso, 'Desligar' volta a ser um shutdown falso.
# ------------------------------------------------------------
Write-Host "`n[3/6] Mantendo o Fast Startup (Inicializacao rapida) DESLIGADO..." -ForegroundColor Cyan
REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power" /V HiberbootEnabled /T REG_DWORD /D 0 /F 2>&1 | Out-Null
Write-Host "  [ok] HiberbootEnabled = 0 (hibernar SIM, shutdown falso NAO)" -ForegroundColor Green

# ------------------------------------------------------------
# 4) Fechar a tampa = HIBERNAR (na tomada e na bateria)
# ------------------------------------------------------------
Write-Host "`n[4/6] Fechar a tampa = hibernar..." -ForegroundColor Cyan
powercfg /setacvalueindex SCHEME_CURRENT $SUB_BUTTONS $GUID_LID   2 2>&1 | Out-Null
powercfg /setdcvalueindex SCHEME_CURRENT $SUB_BUTTONS $GUID_LID   2 2>&1 | Out-Null
# botao de energia segue = desligar (comportamento do setup)
powercfg /setacvalueindex SCHEME_CURRENT $SUB_BUTTONS $GUID_POWER 3 2>&1 | Out-Null
powercfg /setdcvalueindex SCHEME_CURRENT $SUB_BUTTONS $GUID_POWER 3 2>&1 | Out-Null

# rede de seguranca: na bateria, esquecido ABERTO, hiberna em 60 min
powercfg /change hibernate-timeout-dc 60 2>&1 | Out-Null
powercfg /change hibernate-timeout-ac  0 2>&1 | Out-Null

powercfg /setactive SCHEME_CURRENT 2>&1 | Out-Null
Write-Host "  [ok] tampa = hibernar | botao = desligar | bateria: hiberna em 60 min de inatividade" -ForegroundColor Green

# ------------------------------------------------------------
# 5) Mostrar 'Hibernar' no menu Iniciar
# ------------------------------------------------------------
Write-Host "`n[5/6] Colocando 'Hibernar' no menu Iniciar..." -ForegroundColor Cyan
$fly = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FlyoutMenuSettings'
if (-not (Test-Path $fly)) { New-Item -Path $fly -Force | Out-Null }
Set-ItemProperty -Path $fly -Name 'ShowHibernateOption' -Value 1 -Type DWord
Write-Host "  [ok] ShowHibernateOption = 1" -ForegroundColor Green

# ------------------------------------------------------------
# 6) DEPOIS: conferir de verdade (nao confiar na tela)
# ------------------------------------------------------------
Write-Host "`n[6/6] Estados disponiveis DEPOIS:" -ForegroundColor Cyan
$depois = (powercfg /a) -join "`n"
Write-Host $depois -ForegroundColor DarkGray

# valor efetivo da acao da tampa, lido do esquema ativo
$q = (powercfg /query SCHEME_CURRENT $SUB_BUTTONS $GUID_LID 2>&1) -join "`n"
if ($q -notmatch '0x') {
    # em Win11 recente a consulta dirigida pode nao trazer os indices: cai no dump inteiro
    $todo = (powercfg /query 2>&1) -join "`n"
    $pos = $todo.IndexOf($GUID_LID)
    if ($pos -ge 0) { $q = $todo.Substring($pos, [Math]::Min(1200, $todo.Length - $pos)) }
}
$nomes = @{0='Nao fazer nada';1='Suspender';2='Hibernar';3='Desligar';4='Desligar o video'}
# so as duas linhas "Current AC/DC Power Setting Index" tem 0x nesta saida -> serve em
# qualquer idioma do Windows (PT/ES/EN), sem depender do texto traduzido.
$idx = @([regex]::Matches($q, ':\s*0x([0-9a-f]+)'))
$tampaOk = $false
if ($idx.Count -ge 1) {
    $rotulo = @('na tomada', 'na bateria')
    for ($i = 0; $i -lt $idx.Count; $i++) {
        $v = [Convert]::ToInt32($idx[$i].Groups[1].Value, 16)
        $nome = if ($nomes.ContainsKey($v)) { $nomes[$v] } else { "valor $v" }
        $onde = if ($i -lt 2) { $rotulo[$i] } else { '?' }
        if ($v -eq 2) { $tampaOk = $true }
        Write-Host ("  tampa $onde -> $nome") -ForegroundColor Green
    }
} else {
    Write-Host "  [i] nao consegui ler o indice — confira na tela do Painel de Controle" -ForegroundColor Yellow
    Write-Host "  --- saida bruta da consulta (para o TI) ---" -ForegroundColor DarkGray
    Write-Host $q -ForegroundColor DarkGray
}

# O veredito olha o REGISTRO (HibernateEnabled), nao o texto do /a: a palavra
# "Hibernate/Hibernar" aparece TAMBEM na lista de estados NAO disponiveis
# ("Hibernation has not been enabled"), entao procurar a palavra daria falso positivo.
# E nao olha apenas o hiberfil.sys: ele pode demorar a aparecer (falso alarme no 1o uso real).
$hibernaOk = (Test-HibernacaoLigada) -or (Test-Path $hiberFile)
if ($hibernaOk -and $tampaOk) {
    Write-Host "`n=== PRONTO ===" -ForegroundColor Green
    Write-Host "Hibernacao disponivel. Teste agora: feche a tampa, espere ~20s e veja se o PC" -ForegroundColor White
    Write-Host "apaga de verdade (ventoinha e LED param). Abra e o trabalho volta como estava." -ForegroundColor White
} elseif ($hibernaOk) {
    Write-Host "`n=== PRONTO (com uma conferencia) ===" -ForegroundColor Green
    Write-Host "A hibernacao esta LIGADA. Nao consegui LER de volta a acao da tampa" -ForegroundColor White
    Write-Host "(apenas ler; a gravacao do passo 4 nao deu erro) - confira na tela:" -ForegroundColor White
    Write-Host "Abra: Painel de Controle > Opcoes de Energia > 'Escolher o que fazer ao fechar a tampa'" -ForegroundColor White
    Write-Host "e veja se esta em Hibernar nas duas colunas (a opcao ja existe na lista)." -ForegroundColor White
} else {
    Write-Host "`n=== ATENCAO ===" -ForegroundColor Red
    Write-Host "A hibernacao NAO ligou. Motivos possiveis (o proprio /a do passo 6 diz):" -ForegroundColor White
    Write-Host "  - politica de grupo ou firmware bloqueando;" -ForegroundColor White
    Write-Host "  - Hyper-V / VBS ligado (aparece 'o hipervisor nao suporta este estado');" -ForegroundColor White
    Write-Host "  - sem espaco no C: para o hiberfil.sys (~$hiberGB GB)." -ForegroundColor White
    Write-Host "Manda o log do Desktop\UTI-backup\corrigir-hibernacao-log.txt para o TI." -ForegroundColor Yellow
}

Write-Host "`nExtras para o TI (nao mexem em nada):" -ForegroundColor DarkGray
Write-Host "  - quem esta segurando o PC acordado:  powercfg /requests" -ForegroundColor DarkGray
Write-Host "  - quem pode acordar o PC:             powercfg /devicequery wake_armed" -ForegroundColor DarkGray
Write-Host "  - consumo da bateria:                 powercfg /batteryreport /output `"$env:USERPROFILE\bateria.html`"" -ForegroundColor DarkGray
Write-Host "  - se ALGUM DIA quiser o SUSPENDER de volta (traz de volta o LED piscando" -ForegroundColor DarkGray
Write-Host "    depois de desligar):  reg delete `"HKLM\SYSTEM\CurrentControlSet\Control\Power`" /v PlatformAoAcOverride /f  + reboot" -ForegroundColor DarkGray

Stop-Transcript | Out-Null
