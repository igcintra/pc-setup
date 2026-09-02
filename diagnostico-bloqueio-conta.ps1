# ============================================================================
# DIAGNOSTICO DE BLOQUEIO DE CONTA (STATUS_ACCOUNT_LOCKED_OUT) - IG Networks / TI
# ----------------------------------------------------------------------------
# Contexto: desde o build 22528 o Windows 11 vem com politica de bloqueio LIGADA
# de fabrica (10 tentativas erradas -> 10 min trancado), so em instalacao limpa.
# A frota da IGN e' Win 11 Home de fabrica, entao nasceu com isso.
#
# O que ele responde, que e' a unica pergunta que importa:
#   QUEM esta errando a senha — a PESSOA no teclado, ou alguem PELA REDE?
#   - LogonType 2/7/11 = teclado/console  -> chateacao (Caps Lock, layout ES x PT)
#   - LogonType 3/10   = rede/RDP         -> incidente, tem IP de origem
#
# NAO ALTERA NADA. Nao destranca conta, nao mexe em politica. So le e relata.
#
# EXIGE ADMINISTRADOR: o log de Seguranca (evento 4625) so e' legivel elevado.
# Sem admin ele avisa e para — relatorio pela metade nao serve pra decidir.
#
# Uso:  irm https://raw.githubusercontent.com/igcintra/pc-setup/master/diagnostico-bloqueio-conta.ps1 | iex
# ============================================================================

$ErrorActionPreference = 'SilentlyContinue'
$VERSAO = '2026-09-02'
$DIAS   = 14        # janela de busca no log de Seguranca

# ---------------------------------------------------------------- idioma ----
function Get-Idioma {
    if ($env:IGN_IDIOMA -match '^(PT|ES|EN)$') { return $env:IGN_IDIOMA.ToUpper() }
    $c = @()
    try { $c += (Get-UICulture).Name } catch {}
    try { $c += (Get-Culture).Name } catch {}
    try { $c += @((Get-CimInstance Win32_OperatingSystem).MUILanguages) } catch {}
    foreach ($x in ($c | Where-Object { $_ })) {
        if ($x -match '^pt') { return 'PT' }
        if ($x -match '^es') { return 'ES' }
    }
    return 'EN'
}
$IDIOMA = Get-Idioma

$STR = @{
 PT = @{
  titulo='DIAGNOSTICO DE BLOQUEIO DE CONTA - IG Networks'
  intro='Leva alguns segundos, nao altera nada e nao destranca nada.'
  adm_t='PRECISA SER ADMINISTRADOR'
  adm_1='O log de Seguranca do Windows so pode ser lido como administrador.'
  adm_2='NADA foi verificado ainda - nenhum relatorio foi gerado.'
  adm_3='Menu Iniciar > digite PowerShell > botao direito > "Executar como administrador".'
  adm_4='Depois cole de novo o MESMO comando. Sem a senha de admin, avise a TI.'
  fim  ='Pronto. Mande para a TI o arquivo que ficou no seu Desktop:'
  ia   ='Esse arquivo tambem pode ser colado em qualquer IA - ele ja traz as instrucoes dentro.'
 }
 ES = @{
  titulo='DIAGNOSTICO DE BLOQUEO DE CUENTA - IG Networks'
  intro='Toma unos segundos, no cambia nada y no desbloquea nada.'
  adm_t='HACE FALTA SER ADMINISTRADOR'
  adm_1='El registro de Seguridad de Windows solo se puede leer como administrador.'
  adm_2='TODAVIA no se verifico nada - no se genero ningun reporte.'
  adm_3='Menu Inicio > escribe PowerShell > clic derecho > "Ejecutar como administrador".'
  adm_4='Despues pega de nuevo el MISMO comando. Si no tenes la contrasena de admin, avisa a TI.'
  fim  ='Listo. Envia a TI el archivo que quedo en tu Escritorio:'
  ia   ='Este archivo tambien se puede pegar en cualquier IA - ya trae las instrucciones dentro.'
 }
 EN = @{
  titulo='ACCOUNT LOCKOUT DIAGNOSTIC - IG Networks'
  intro='Takes a few seconds, changes nothing and unlocks nothing.'
  adm_t='ADMINISTRATOR REQUIRED'
  adm_1='The Windows Security log can only be read as administrator.'
  adm_2='NOTHING has been checked yet - no report was generated.'
  adm_3='Start menu > type PowerShell > right-click > "Run as administrator".'
  adm_4='Then paste the SAME command again. If you do not have the admin password, tell IT.'
  fim  ='Done. Please send IT the file left on your Desktop:'
  ia   ='This file can also be pasted into any AI - the instructions are inside it.'
 }
}[$IDIOMA]

Write-Host ""
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host ("  " + $STR.titulo) -ForegroundColor Cyan
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host ""
Write-Host $STR.intro -ForegroundColor Gray
Write-Host ""

# --------------------------------------------------- exigencia de admin ----
# Mesmo desenho do rastreador (v2026-09-01): #Requires nao vale em irm|iex e
# exit fecharia a janela antes de a pessoa ler.
$EH_ADMIN = $false
try {
    $EH_ADMIN = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
                ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} catch { $EH_ADMIN = $false }

if (-not $EH_ADMIN) {
    Write-Host ("=" * 60) -ForegroundColor Yellow
    Write-Host ("  " + $STR.adm_t) -ForegroundColor Yellow
    Write-Host ("=" * 60) -ForegroundColor Yellow
    Write-Host ""
    Write-Host $STR.adm_1 -ForegroundColor Gray
    Write-Host $STR.adm_2 -ForegroundColor White
    Write-Host ""
    Write-Host $STR.adm_3 -ForegroundColor Cyan
    Write-Host $STR.adm_4 -ForegroundColor Cyan
    Write-Host ""
    return
}

function Safe([scriptblock]$sb, $padrao = 'n/d') {
    try { $v = & $sb; if ($null -eq $v -or "$v" -eq '') { return $padrao }; return $v }
    catch { return $padrao }
}

$R = New-Object System.Collections.ArrayList
function W($t) { [void]$R.Add([string]$t) }
$limites = New-Object System.Collections.ArrayList

# ============================ dicionarios ===================================
$TIPO_LOGON = @{
  2='INTERATIVO (teclado da maquina)'; 3='REDE (SMB/compartilhamento)'
  4='TAREFA AGENDADA'; 5='SERVICO'; 7='DESBLOQUEIO DE TELA'
  8='REDE EM TEXTO CLARO'; 9='NOVAS CREDENCIAIS (runas)'
  10='REMOTO / RDP'; 11='CACHE (interativo sem contato com o dominio)'
}
# Este e' o campo que diz o MOTIVO real da falha. Sem ele, "senha errada" e
# "usuario nao existe" viram a mesma linha no relatorio.
$SUBSTATUS = @{
  '0xc000006a'='SENHA ERRADA'
  '0xc0000064'='USUARIO NAO EXISTE (nome digitado errado ou conta apagada)'
  '0xc0000234'='CONTA JA ESTAVA TRANCADA (tentativa durante o bloqueio)'
  '0xc0000072'='CONTA DESATIVADA'
  '0xc0000071'='SENHA EXPIRADA'
  '0xc0000193'='CONTA EXPIRADA'
  '0xc000006f'='FORA DO HORARIO PERMITIDO'
  '0xc0000070'='ESTACAO NAO PERMITIDA'
  '0xc0000224'='PRECISA TROCAR A SENHA NO PROXIMO LOGON'
  '0xc0000133'='RELOGIO FORA DE HORA'
}
$TECLADO = @{
  '00000416'='Portugues (Brasil ABNT)'; '00010416'='Portugues (Brasil ABNT2)'
  '00000816'='Portugues (Portugal)';    '0000080a'='Espanhol (Mexico)'
  '0000040a'='Espanhol (Espanha)';      '00000c0a'='Espanhol (Espanha moderno)'
  '00000409'='Ingles (EUA)';            '00000809'='Ingles (Reino Unido)'
  '00080409'='Ingles (EUA internacional)'
}

Write-Host "[1/5] Lendo a politica de bloqueio..." -ForegroundColor White

# ============================ 1. politica ===================================
$pol = [ordered]@{ limite='n/d'; duracao_min='n/d'; janela_min='n/d' }
try {
    $na = (net accounts) 2>$null
    foreach ($l in $na) {
        if ($l -match '(?i)(lockout threshold|limite de bloqueio|umbral de bloqueo)\D*(\d+|Never|Nunca)') { $pol.limite      = $Matches[2] }
        if ($l -match '(?i)(lockout duration|dura(c|ç)(a|ã)o do bloqueio|duraci(o|ó)n del bloqueo)\D*(\d+)') { $pol.duracao_min = $Matches[4] }
        if ($l -match '(?i)(observation window|janela de observa|ventana de observaci)\D*(\d+)')             { $pol.janela_min  = $Matches[2] }
    }
} catch { [void]$limites.Add('Nao foi possivel ler "net accounts".') }

Write-Host "[2/5] Vendo o estado das contas locais..." -ForegroundColor White

# ============================ 2. contas =====================================
$contas = @()
try {
    foreach ($u in (Get-LocalUser)) {
        $trancada = 'n/d'
        try {
            $adsi = [ADSI]("WinNT://./" + $u.Name + ",user")
            if ($null -ne $adsi.IsAccountLocked) { $trancada = [bool]$adsi.IsAccountLocked.Value }
        } catch {}
        $contas += [pscustomobject]@{
            nome            = $u.Name
            habilitada      = [bool]$u.Enabled
            trancada_agora  = $trancada
            ultimo_logon    = Safe { if ($u.LastLogon) { $u.LastLogon.ToString('yyyy-MM-dd HH:mm') } }
            senha_alterada  = Safe { if ($u.PasswordLastSet) { $u.PasswordLastSet.ToString('yyyy-MM-dd HH:mm') } }
        }
    }
} catch { [void]$limites.Add('Nao foi possivel listar as contas locais (Get-LocalUser).') }

Write-Host "[3/5] Conferindo se a auditoria de logon esta ligada..." -ForegroundColor White

# ============================ 3. auditoria ==================================
# Se a auditoria de falha estiver DESLIGADA, o log fica vazio e "sem eventos"
# NAO quer dizer "ninguem tentou". Sem isso o relatorio enganaria.
$auditoria = 'n/d'
try {
    $ap = (auditpol /get /subcategory:"Logon" /r) 2>$null
    if ($ap) { $auditoria = (($ap | Select-Object -Skip 1) -join ' ') }
} catch {}
if ($auditoria -match '(?i)no auditing|sem auditoria|sin auditor') {
    [void]$limites.Add('AUDITORIA DE LOGON DESLIGADA: o log de falhas pode estar vazio por configuracao, nao por ausencia de tentativa.')
}

Write-Host "[4/5] Lendo o log de Seguranca (pode demorar um pouco)..." -ForegroundColor White

# ============================ 4. eventos ====================================
$desde   = (Get-Date).AddDays(-$DIAS)
$falhas  = @()
$locks   = @()

try {
    $ev4625 = Get-WinEvent -FilterHashtable @{LogName='Security'; Id=4625; StartTime=$desde} -ErrorAction Stop
} catch { $ev4625 = @() }
try {
    $ev4740 = Get-WinEvent -FilterHashtable @{LogName='Security'; Id=4740; StartTime=$desde} -ErrorAction Stop
} catch { $ev4740 = @() }

if (-not $ev4625 -and -not $ev4740) {
    [void]$limites.Add("Nenhum evento 4625/4740 nos ultimos $DIAS dias. Pode ser log rotacionado (log pequeno) ou auditoria desligada.")
}

# Le por NOME do campo (XML), nao por indice: indice muda entre versoes do Windows.
function Campos($ev) {
    $h = @{}
    try {
        $x = [xml]$ev.ToXml()
        foreach ($d in $x.Event.EventData.Data) { $h[$d.Name] = [string]$d.'#text' }
    } catch {}
    return $h
}

foreach ($e in $ev4625) {
    $c = Campos $e
    $sub = ([string]$c['SubStatus']).ToLower()
    $lt  = [string]$c['LogonType']
    $ip  = [string]$c['IpAddress']
    if ($ip -in @('-','::1','127.0.0.1','')) { $ip = 'local' }
    $falhas += [pscustomobject]@{
        quando  = $e.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
        conta   = [string]$c['TargetUserName']
        tipo    = $lt
        tipo_txt= $(if ($TIPO_LOGON.ContainsKey([int]($lt -as [int]))) { $TIPO_LOGON[[int]$lt] } else { "tipo $lt" })
        motivo  = $(if ($SUBSTATUS.ContainsKey($sub)) { $SUBSTATUS[$sub] } else { "codigo $sub" })
        origem  = $ip
        estacao = [string]$c['WorkstationName']
        processo= [string]$c['ProcessName']
    }
}
foreach ($e in $ev4740) {
    $c = Campos $e
    $locks += [pscustomobject]@{
        quando  = $e.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
        conta   = [string]$c['TargetUserName']
        origem  = [string]$c['TargetDomainName']
    }
}

Write-Host "[5/5] Montando o relatorio..." -ForegroundColor White

# ============================ 5. veredito ===================================
$de_rede    = @($falhas | Where-Object { $_.tipo -in @('3','8','10') })
$de_teclado = @($falhas | Where-Object { $_.tipo -in @('2','7','11') })
$ips_ext    = @($de_rede | Where-Object { $_.origem -ne 'local' } | Select-Object -Expand origem -Unique)

if ($falhas.Count -eq 0) {
    $veredito = 'SEM EVENTOS'
} elseif ($de_rede.Count -gt $de_teclado.Count -and $ips_ext.Count -gt 0) {
    $veredito = 'REDE'
} elseif ($de_rede.Count -gt 0 -and $ips_ext.Count -gt 0) {
    $veredito = 'MISTO'
} else {
    $veredito = 'TECLADO'
}

# ============================ relatorio =====================================
$desktop = [Environment]::GetFolderPath('Desktop')
if (-not $desktop -or -not (Test-Path $desktop)) { $desktop = "$env:USERPROFILE\Desktop" }
$limpaU  = ($env:USERNAME     -replace '[^A-Za-z0-9._-]','')
$limpaH  = ($env:COMPUTERNAME -replace '[^A-Za-z0-9._-]','')
$nomeArq = "bloqueio-conta-$limpaU-$limpaH-$(Get-Date -Format 'yyyy-MM-dd').txt"
$arq     = Join-Path $desktop $nomeArq

W ("=" * 78)
W " DIAGNOSTICO DE BLOQUEIO DE CONTA - IG Networks"
W ("=" * 78)
W ""
W "### PARA QUEM RECEBEU ESTE ARQUIVO"
W "PT: Resultado de uma verificacao. Nada foi alterado. Envie o arquivo para a TI."
W "ES: Resultado de una verificacion. No se cambio nada. Envia el archivo a TI."
W ""
W "### INSTRUCOES PARA UMA IA (copie este arquivo inteiro no chat)"
W "Estacao Windows de uma empresa de telecom (IG Networks, LATAM), FORA DE DOMINIO."
W "A pessoa viu na tela: 'A conta referenciada esta bloqueada'. Desde o build 22528 o"
W "Windows 11 tranca a conta sozinho apos N senhas erradas - isso e' padrao de fabrica,"
W "nao configuracao da empresa."
W "Responda nesta ordem:"
W "  1. VEREDITO em uma frase, para quem nao e tecnico."
W "  2. QUEM estava errando a senha: a pessoa no teclado (LogonType 2/7/11) ou alguem"
W "     pela rede (LogonType 3/10, com IP). Essa e' a pergunta que decide tudo."
W "  3. Se for teclado: olhe [TECLADOS] - layout ES x PT troca de lugar os simbolos e a"
W "     senha 'certa' vira errada sempre. Veja tambem se o motivo e' USUARIO NAO EXISTE"
W "     (nome errado), que nao e' problema de senha."
W "  4. Se houver IP de origem em [FALHAS], trate como tentativa de invasao ate provar o"
W "     contrario, e NAO recomende desligar a politica de bloqueio."
W "  5. Repita a secao [COBERTURA-E-LIMITES] - ela e' obrigatoria na sua resposta."
W "Nao invente evento que nao esta no arquivo. Dados estruturados em [DADOS-JSON]."
W ""
W ("=" * 78)
W ""
W "[MAQUINA]"
W ("  gerado_em      = " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
W ("  computador     = " + (Safe { $env:COMPUTERNAME }))
W ("  usuario_que_rodou = " + (Safe { $env:USERNAME }))
W ("  dominio        = " + (Safe { $env:USERDOMAIN }))
W ("  windows        = " + (Safe { (Get-CimInstance Win32_OperatingSystem).Caption }))
W ("  build          = " + (Safe { (Get-CimInstance Win32_OperatingSystem).BuildNumber }))
W ("  idioma_windows = " + (Safe { (Get-UICulture).Name }))
W ("  versao_script  = " + $VERSAO)
W ""
W "[POLITICA DE BLOQUEIO]"
W ("  limite de tentativas = " + $pol.limite)
W ("  duracao do bloqueio  = " + $pol.duracao_min + " min")
W ("  janela de contagem   = " + $pol.janela_min + " min")
W "  (padrao de fabrica do Win 11 em instalacao limpa: 10 tentativas / 10 min)"
W ""
W "[TECLADOS INSTALADOS]"
W "  Layout diferente do esperado troca de lugar os simbolos: a pessoa digita a senha"
W "  certa e o Windows recebe outra coisa. Causa classica em frota PT + ES."
try {
    $pre = Get-ItemProperty 'HKCU:\Keyboard Layout\Preload' -ErrorAction Stop
    foreach ($p in $pre.PSObject.Properties) {
        if ($p.Name -match '^\d+$') {
            $cod = $p.Value.ToString().ToLower()
            $nom = if ($TECLADO.ContainsKey($cod)) { $TECLADO[$cod] } else { "codigo $cod" }
            W ("  - " + $nom)
        }
    }
} catch { W "  (nao foi possivel ler os layouts do usuario atual)" }
W ""
W "[CONTAS LOCAIS]"
foreach ($c in $contas) {
    W ("  - {0,-22} habilitada={1,-5} trancada_agora={2,-5} ultimo_logon={3}" -f `
        $c.nome, $c.habilitada, $c.trancada_agora, $c.ultimo_logon)
}
W ""
W "[AUDITORIA DE LOGON]"
W ("  " + $auditoria)
W ""
W ("[FALHAS DE LOGON - ultimos $DIAS dias]")
W ("  total = " + $falhas.Count + " | teclado/console = " + $de_teclado.Count + " | rede/RDP = " + $de_rede.Count)
if ($ips_ext.Count) { W ("  IPs de origem vistos: " + ($ips_ext -join ', ')) }
W ""
if ($falhas.Count) {
    W "  Por conta e motivo:"
    foreach ($g in ($falhas | Group-Object conta, motivo | Sort-Object Count -Descending)) {
        W ("    {0,-5}x  {1}" -f $g.Count, $g.Name)
    }
    W ""
    W "  Ultimas 40 tentativas (mais recente primeiro):"
    W "  DATA/HORA            CONTA            TIPO                              MOTIVO / ORIGEM"
    foreach ($f in ($falhas | Sort-Object quando -Descending | Select-Object -First 40)) {
        W ("  {0}  {1,-15}  {2,-32}  {3} [{4}]" -f $f.quando, $f.conta, $f.tipo_txt, $f.motivo, $f.origem)
    }
} else {
    W "  Nenhuma falha registrada na janela."
}
W ""
W "[BLOQUEIOS REGISTRADOS - evento 4740]"
if ($locks.Count) {
    foreach ($l in $locks) { W ("  " + $l.quando + "  conta=" + $l.conta + "  origem=" + $l.origem) }
} else {
    W "  Nenhum evento 4740. Em maquina fora de dominio ele frequentemente NAO e' gravado"
    W "  mesmo havendo bloqueio - a prova util aqui sao as falhas 4625 acima."
}
W ""
W "[VEREDITO]"
W ("  resultado = " + $veredito)
W "  TECLADO     = as falhas vem do console: senha digitada errada, Caps Lock ou layout"
W "  REDE        = a maioria vem da rede/RDP, com IP: tratar como tentativa de invasao"
W "  MISTO       = as duas coisas; olhar os IPs antes de concluir"
W "  SEM EVENTOS = nada no log; ler [COBERTURA-E-LIMITES] antes de dizer que esta tudo bem"
W ""
W "[COBERTURA-E-LIMITES]"
W "  - Le o log de Seguranca dos ultimos $DIAS dias. Log pequeno rotaciona e apaga o que"
W "    interessa: 'sem eventos' pode ser log curto, nao ausencia de tentativa."
W "  - Se a auditoria de falha estiver desligada, NAO existe registro para ler."
W "  - O Windows nao registra Caps Lock nem qual tecla foi digitada. O layout instalado e'"
W "    indicio, nao prova."
W "  - Bloqueio por senha errada e ataque de forca bruta produzem o MESMO evento. O que"
W "    separa os dois e' o tipo de logon e o IP de origem."
W "  - Este script nao destranca conta e nao muda politica. Com o limite padrao, a conta"
W "    destranca sozinha ao fim da duracao do bloqueio."
if ($limites.Count) { W ""; foreach ($l in $limites) { W ("  * " + $l) } }
W ""
W "[DADOS-JSON]"
W ((([pscustomobject]@{
      maquina   = [ordered]@{
                    computador = "$env:COMPUTERNAME"; usuario_que_rodou = "$env:USERNAME"
                    dominio = "$env:USERDOMAIN"; versao_script = $VERSAO
                    gerado_em = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                  }
      politica  = $pol
      contas    = @($contas)
      falhas    = @($falhas)
      bloqueios = @($locks)
      veredito  = $veredito
      limites   = @($limites)
   }) | ConvertTo-Json -Depth 6))
W ""
W ("=" * 78)
W (" Gerado por diagnostico-bloqueio-conta (IG Networks / TI) - versao " + $VERSAO)
W ("=" * 78)

try {
    $R | Set-Content -Path $arq -Encoding UTF8 -ErrorAction Stop
} catch {
    $arq = Join-Path $env:USERPROFILE $nomeArq
    try { $R | Set-Content -Path $arq -Encoding UTF8 -ErrorAction Stop }
    catch { $arq = Join-Path $env:TEMP $nomeArq; $R | Set-Content -Path $arq -Encoding UTF8 }
}

Write-Host ""
Write-Host ("=" * 60) -ForegroundColor Green
Write-Host ("  " + $STR.fim) -ForegroundColor Green
Write-Host ("=" * 60) -ForegroundColor Green
Write-Host ("  " + $nomeArq) -ForegroundColor White
Write-Host ""
Write-Host ("VEREDITO: " + $veredito) -ForegroundColor Yellow
Write-Host ("Falhas na janela: " + $falhas.Count + "  (teclado " + $de_teclado.Count + " / rede " + $de_rede.Count + ")") -ForegroundColor Gray
Write-Host ""
Write-Host $STR.ia -ForegroundColor DarkGray
Write-Host ""
