# ============================================================================
# reteste-wifi.ps1 - RETESTE da placa Wi-Fi que "conecta e cai" - IG Networks
# ----------------------------------------------------------------------------
# Uso (PowerShell comum - ele se AUTO-ELEVA, so clicar "Sim" no UAC):
#   irm https://raw.githubusercontent.com/igcintra/pc-setup/master/reteste-wifi.ps1 | iex
#
# Para que serve: refazer, de uma vez so, os testes que FORAM UTEIS no caso da
# aamorin (Lenovo IdeaPad 3 15ITL6, Intel AC 9560) e responder UMA pergunta:
#   >>> O PROBLEMA CONTINUA OU NAO? <<<
# Faz:
#   [1] Placa/driver + se o fix de energia ainda esta aplicado (reinstalar
#       driver RESETA o fix - por isso confere antes de julgar).
#   [2] Historico de quedas: eventos 8003 (WLAN) e Netwtw* 6062 (driver resetando).
#   [3] TESTE AO VIVO: ping em rajada no ROTEADOR e na INTERNET (separa "a placa
#       caiu" de "a internet caiu") + vigia o status da interface e o sinal.
#   [4] Eventos que aconteceram DURANTE o teste (prova de queda sem ninguem mexer).
#   [5] VEREDITO em portugues, no topo do log e na tela.
# Log: Desktop\reteste-wifi-AAAA-MM-DD_HHmm.txt (e ja vai pro clipboard).
# Dica: rode 1x na rede do escritorio/casa e 1x no HOTSPOT DO CELULAR. Cair nas
# duas = placa/driver; cair so numa = roteador. Foi assim que a maquina da Anna
# foi condenada em 11/08/2026. Ver memoria [[reference_wifi_conecta_e_cai]].
# ============================================================================

$ErrorActionPreference = 'Continue'
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Elevando... clique SIM na janela que vai abrir (o teste roda la)." -ForegroundColor Yellow
    # leva junto o Desktop DESTE usuario (resolve OneDrive) p/ o log nao cair em outro perfil
    $desk = [Environment]::GetFolderPath('Desktop')
    $cmd = "`$env:RETESTE_DESK='$desk'; irm https://raw.githubusercontent.com/igcintra/pc-setup/master/reteste-wifi.ps1 | iex"
    Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-Command',$cmd
    return
}

# ---------------------------------------------------------------- duracao
Write-Host ""
Write-Host "=== RETESTE DA PLACA WI-FI - IG Networks ===" -ForegroundColor Cyan
Write-Host "Vou monitorar a conexao por alguns minutos e dizer se o problema continua."
Write-Host "NAO precisa mexer no PC durante o teste (pode deixar rodando)." -ForegroundColor DarkGray
$resp = Read-Host "Quantos MINUTOS de teste? [Enter = 10]"
$mins = 10
if ($resp -and ($resp -as [int]) -and ([int]$resp -ge 1)) { $mins = [int]$resp }

$inicio = Get-Date
$R = New-Object System.Collections.ArrayList
function Add-L([string]$t) { [void]$R.Add($t) }

Add-L "=== RETESTE WI-FI - $($inicio.ToString('dd/MM/yyyy HH:mm')) - $env:COMPUTERNAME ==="
try {
    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
    $bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
    Add-L ("Maquina : {0} {1} | Serial: {2} | Windows {3} build {4}" -f $cs.Manufacturer, $cs.Model, $bios.SerialNumber, (Get-CimInstance Win32_OperatingSystem).Caption, [System.Environment]::OSVersion.Version.Build)
} catch {}
Add-L "Duracao do teste: $mins minuto(s)"
Add-L ""

# =========================================================== [1] PLACA/FIX
Add-L "--- [1] PLACA, DRIVER E FIX DE ENERGIA ---"
$wifi = Get-NetAdapter | Where-Object { $_.InterfaceDescription -match 'Wireless|Wi-?Fi|WLAN' } | Select-Object -First 1
if (-not $wifi) {
    Add-L "ERRO CRITICO: nenhum adaptador Wi-Fi encontrado no sistema."
    Add-L "Isso ja e' um achado: a placa SUMIU do Windows (ver Gerenciador de Dispositivos)."
    $semPlaca = $true
} else {
    $semPlaca = $false
    Add-L ("Placa   : {0}" -f $wifi.InterfaceDescription)
    Add-L ("Status  : {0} | Velocidade: {1}" -f $wifi.Status, $wifi.LinkSpeed)
    Add-L ("Driver  : versao {0} de {1} ({2})" -f $wifi.DriverVersion, $wifi.DriverDate, $wifi.DriverProvider)
    $pnp = Get-PnpDevice -FriendlyName $wifi.InterfaceDescription -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($pnp) { Add-L ("PnP     : {0} | erro/status: {1}" -f $pnp.InstanceId, $pnp.Status) }
}

# fix de energia (MSPower) - Enable=False e' o desejado
$energiaOK = $null
if (-not $semPlaca -and $pnp) {
    $prefixo = ($pnp.InstanceId.Split('\')[0..1] -join '\')
    $alvo = Get-CimInstance -Namespace root\wmi -ClassName MSPower_DeviceEnable -ErrorAction SilentlyContinue |
            Where-Object { $_.InstanceName -like "$prefixo*" } | Select-Object -First 1
    if ($alvo) {
        $energiaOK = (-not $alvo.Enable)
        if ($energiaOK) { Add-L "Economia de energia da placa: DESLIGADA (fix aplicado, OK)" }
        else            { Add-L "Economia de energia da placa: LIGADA  <-- FIX PERDIDO (reinstalacao de driver reseta)" }
    } else { Add-L "Economia de energia: nao consegui casar o MSPower dessa placa (inconclusivo)" }
}
if (-not $semPlaca) {
    $mimo = Get-NetAdapterAdvancedProperty -Name $wifi.Name -ErrorAction SilentlyContinue | Where-Object DisplayName -match 'MIMO'
    foreach ($m in $mimo) { Add-L ("MIMO    : {0} = {1}" -f $m.DisplayName, $m.DisplayValue) }
    Add-L ""
    Add-L "--- SINAL / REDE ATUAL ---"
    Add-L ((netsh wlan show interfaces) -join "`r`n")
}
Add-L ""

# ==================================================== [2] HISTORICO QUEDAS
Add-L "--- [2] HISTORICO DE QUEDAS (ultimos 7 dias) ---"
$desde = (Get-Date).AddDays(-7)
$ev8003 = @(Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-WLAN-AutoConfig/Operational';Id=8003;StartTime=$desde} -ErrorAction SilentlyContinue)
$evDrv  = @(Get-WinEvent -FilterHashtable @{LogName='System';StartTime=$desde} -ErrorAction SilentlyContinue | Where-Object { $_.ProviderName -match 'Netwtw|netwlv|iaLPSS|NDIS' })
Add-L ("Quedas de Wi-Fi (evento 8003)         : {0}" -f $ev8003.Count)
Add-L ("Eventos do driver Intel (Netwtw/NDIS) : {0}  (6062 em serie = driver resetando)" -f $evDrv.Count)
if ($ev8003.Count) {
    Add-L "Ultimas quedas:"
    $ev8003 | Select-Object -First 8 | ForEach-Object {
        $motivo = ($_.Message -split "`r?`n" | Select-String 'Raz|Reason|Motivo' | Select-Object -First 1)
        Add-L ("  {0}  {1}" -f $_.TimeCreated, ($motivo -replace '\s+',' '))
    }
}
if ($evDrv.Count) {
    Add-L "Eventos de driver (id / hora):"
    $evDrv | Group-Object Id | Sort-Object Count -Descending | Select-Object -First 5 | ForEach-Object {
        Add-L ("  id {0} x{1}  (ultimo: {2})" -f $_.Name, $_.Count, ($_.Group | Select-Object -First 1).TimeCreated)
    }
}
Add-L ""

# ======================================================= [3] TESTE AO VIVO
Add-L "--- [3] TESTE AO VIVO ($mins min) ---"
$amostras = 0; $falhaGw = 0; $falhaNet = 0
$maiorRajada = 0; $rajadas = 0; $quedasIface = 0
$sinais = New-Object System.Collections.ArrayList
$linhaTempo = New-Object System.Collections.ArrayList
$perdaGw = 0; $perdaNet = 0; $sinalMed = $null
if ($semPlaca) {
    Add-L "PULADO: nao ha placa Wi-Fi no sistema para testar."
} else {
    $gw = (Get-NetIPConfiguration -ErrorAction SilentlyContinue | Where-Object { $_.NetAdapter.Status -eq 'Up' -and $_.IPv4DefaultGateway } | Select-Object -First 1).IPv4DefaultGateway.NextHop
    if (-not $gw) { $gw = '' }
    Add-L ("Roteador (gateway): {0} | Internet: 8.8.8.8" -f $(if ($gw) { $gw } else { 'NAO ENCONTRADO' }))

    $ping = New-Object System.Net.NetworkInformation.Ping
    $fim = $inicio.AddMinutes($mins)
    $rajadaAtual = 0; $emRajada = $false

    Write-Host ""
    Write-Host "Testando... (deixe esta janela aberta; termina em $mins min)" -ForegroundColor Yellow
    Write-Host "Legenda: . = ok   ! = falhou no roteador   x = falhou so na internet   D = placa desconectou" -ForegroundColor DarkGray

    while ((Get-Date) -lt $fim) {
        $amostras++
        $okGw = $true; $okNet = $true
        if ($gw)  { try { $okGw  = ($ping.Send($gw, 1000).Status -eq 'Success') }      catch { $okGw = $false } }
        try       { $okNet = ($ping.Send('8.8.8.8', 1500).Status -eq 'Success') }      catch { $okNet = $false }

        $st = (Get-NetAdapter -Name $wifi.Name -ErrorAction SilentlyContinue).Status
        if ($st -and $st -ne 'Up') {
            $quedasIface++
            [void]$linhaTempo.Add(("{0}  PLACA {1}" -f (Get-Date -Format 'HH:mm:ss'), $st))
            Write-Host "D" -NoNewline -ForegroundColor Red
        }

        if (-not $okGw)  { $falhaGw++ }
        if (-not $okNet) { $falhaNet++ }

        if ((-not $okGw) -or ($gw -eq '' -and -not $okNet)) {
            $rajadaAtual++
            if ($rajadaAtual -gt $maiorRajada) { $maiorRajada = $rajadaAtual }
            if ($rajadaAtual -eq 3 -and -not $emRajada) {
                $rajadas++; $emRajada = $true
                [void]$linhaTempo.Add(("{0}  RAJADA de falhas comecou (roteador inalcancavel)" -f (Get-Date -Format 'HH:mm:ss')))
            }
            Write-Host "!" -NoNewline -ForegroundColor Red
        } elseif (-not $okNet) {
            $rajadaAtual = 0; $emRajada = $false
            Write-Host "x" -NoNewline -ForegroundColor Yellow
        } else {
            $rajadaAtual = 0; $emRajada = $false
            Write-Host "." -NoNewline -ForegroundColor DarkGray
        }

        if ($amostras % 15 -eq 0) {
            $sig = (netsh wlan show interfaces | Select-String 'Sinal|Signal' | Select-Object -First 1)
            if ($sig) {
                $num = ([regex]::Match($sig.ToString(), '(\d+)%')).Groups[1].Value
                if ($num) { [void]$sinais.Add([int]$num) }
            }
            Write-Host "" # quebra de linha a cada ~30s
        }
        Start-Sleep -Seconds 2
    }
    Write-Host ""

    $perdaGw  = if ($amostras) { [math]::Round(100 * $falhaGw  / $amostras, 1) } else { 0 }
    $perdaNet = if ($amostras) { [math]::Round(100 * $falhaNet / $amostras, 1) } else { 0 }
    $sinalMed = if ($sinais.Count) { [math]::Round(($sinais | Measure-Object -Average).Average) } else { $null }

    Add-L ("Amostras (1 a cada 2s)      : {0}" -f $amostras)
    Add-L ("Falhas ate o ROTEADOR       : {0}  ({1}%)" -f $falhaGw, $perdaGw)
    Add-L ("Falhas ate a INTERNET       : {0}  ({1}%)" -f $falhaNet, $perdaNet)
    Add-L ("Rajadas (3+ falhas seguidas): {0}  | maior sequencia: {1} falhas seguidas" -f $rajadas, $maiorRajada)
    Add-L ("Placa saiu do ar (status != Up): {0} vez(es)" -f $quedasIface)
    if ($sinalMed -ne $null) { Add-L ("Sinal medio durante o teste : {0}%" -f $sinalMed) }
    if ($linhaTempo.Count) { Add-L "Linha do tempo:"; $linhaTempo | ForEach-Object { Add-L ("  " + $_) } }
}
Add-L ""

# ============================================ [4] EVENTOS DURANTE O TESTE
Add-L "--- [4] EVENTOS DURANTE O TESTE (prova de queda sem ninguem mexer) ---"
$novos8003 = @(Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-WLAN-AutoConfig/Operational';Id=8003;StartTime=$inicio} -ErrorAction SilentlyContinue)
$novosDrv  = @(Get-WinEvent -FilterHashtable @{LogName='System';StartTime=$inicio} -ErrorAction SilentlyContinue | Where-Object { $_.ProviderName -match 'Netwtw|netwlv|NDIS' })
Add-L ("Quedas 8003 durante o teste        : {0}" -f $novos8003.Count)
Add-L ("Eventos de driver durante o teste  : {0}" -f $novosDrv.Count)
$novos8003 | ForEach-Object { Add-L ("  8003 {0}" -f $_.TimeCreated) }
$novosDrv  | ForEach-Object { Add-L ("  {0} id {1} {2}" -f $_.ProviderName, $_.Id, $_.TimeCreated) }
Add-L ""

# ==================================================== [5] VEREDITO
$culpaPlaca = ($quedasIface -gt 0) -or ($novos8003.Count -gt 0) -or ($novosDrv.Count -gt 0) -or ($rajadas -gt 0) -or ($perdaGw -ge 2)
$soInternet = (-not $culpaPlaca) -and ($perdaNet -ge 2)

$V = New-Object System.Collections.ArrayList
[void]$V.Add("############################################################")
if ($semPlaca) {
    [void]$V.Add("# VEREDITO: O PROBLEMA CONTINUA - E ESTA PIOR")
    [void]$V.Add("# A placa Wi-Fi nem aparece mais no Windows.")
    [void]$V.Add("# Proximo passo: Gerenciador de Dispositivos (procurar item com !)")
    [void]$V.Add("# e reencaixe do modulo M.2 + antenas na bancada.")
} elseif ($culpaPlaca) {
    [void]$V.Add("# VEREDITO: O PROBLEMA CONTINUA. A placa segue caindo.")
    [void]$V.Add("#")
    [void]$V.Add("# Provas:")
    if ($quedasIface -gt 0)     { [void]$V.Add("#  - a placa saiu do ar $quedasIface vez(es) DURANTE o teste") }
    if ($novos8003.Count -gt 0) { [void]$V.Add("#  - $($novos8003.Count) queda(s) de Wi-Fi (evento 8003) durante o teste") }
    if ($novosDrv.Count -gt 0)  { [void]$V.Add("#  - $($novosDrv.Count) evento(s) do driver Intel durante o teste (driver resetando)") }
    if ($rajadas -gt 0)         { [void]$V.Add("#  - $rajadas rajada(s) de falha ate o ROTEADOR (maior: $maiorRajada seguidas)") }
    if ($perdaGw -ge 2)         { [void]$V.Add("#  - $perdaGw% de perda ate o proprio roteador (nao e' a internet)") }
    [void]$V.Add("#")
    if ($energiaOK -eq $false) {
        [void]$V.Add("# ATENCAO ANTES DE CONDENAR: o fix de energia esta PERDIDO.")
        [void]$V.Add("# Rode fix-wifi-energia.ps1 e repita ESTE teste antes de decidir.")
    } else {
        [void]$V.Add("# Fix de energia OK e mesmo assim cai => sobra HARDWARE.")
        [void]$V.Add("# Sem placa sobressalente: a AC 9560 e' CNVi (metade no chipset),")
        [void]$V.Add("# so serve modulo CNVi compativel - transplante as cegas nao vale.")
        [void]$V.Add("# Caminho: reencaixar modulo M.2 + 2 cabinhos da antena e retestar;")
        [void]$V.Add("# se persistir, a maquina esta condenada para uso em rede sem fio")
        [void]$V.Add("# (alternativa barata: adaptador Wi-Fi USB).")
    }
} elseif ($soInternet) {
    [void]$V.Add("# VEREDITO: A PLACA SE COMPORTOU BEM neste teste.")
    [void]$V.Add("# O roteador respondeu sempre; quem falhou foi a INTERNET ($perdaNet%).")
    [void]$V.Add("# Isso aponta para link/roteador, nao para a placa.")
} else {
    [void]$V.Add("# VEREDITO: NAO REPRODUZIU nesta janela de $mins minuto(s).")
    [void]$V.Add("# Zero quedas, zero eventos, perda desprezivel.")
    [void]$V.Add("# ISSO NAO ABSOLVE A PLACA: o defeito e' intermitente.")
    [void]$V.Add("# Rode de novo por 60 min e TAMBEM no hotspot do celular.")
}
if ($sinalMed -ne $null -and $sinalMed -lt 50) {
    [void]$V.Add("#")
    [void]$V.Add("# OBS: sinal medio $sinalMed% (fraco) - chegue mais perto do roteador")
    [void]$V.Add("# para o teste nao confundir 'placa ruim' com 'longe demais'.")
}
[void]$V.Add("############################################################")

$vTexto = $V -join "`r`n"
$cabecalho = $R[0]
$corpo = ($R | Select-Object -Skip 1) -join "`r`n"
$final = $cabecalho + "`r`n`r`n" + $vTexto + "`r`n`r`n" + $corpo + "`r`n`r`n" + $vTexto

$stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
$desktop = if ($env:RETESTE_DESK -and (Test-Path $env:RETESTE_DESK)) { $env:RETESTE_DESK } else { [Environment]::GetFolderPath('Desktop') }
if (-not $desktop -or -not (Test-Path $desktop)) { $desktop = "$env:USERPROFILE\Desktop" }
$out = Join-Path $desktop "reteste-wifi-$stamp.txt"
$final | Out-File $out -Encoding utf8
try { Get-Content $out -Raw | Set-Clipboard } catch {}

Write-Host ""
$cor = if ($semPlaca -or $culpaPlaca) { 'Red' } elseif ($soInternet) { 'Yellow' } else { 'Green' }
Write-Host $vTexto -ForegroundColor $cor
Write-Host ""
Write-Host "Log completo: $out" -ForegroundColor Cyan
Write-Host "(ja copiado para a area de transferencia - so colar no Slack com Ctrl+V)" -ForegroundColor DarkGray
Read-Host "Enter para fechar"
