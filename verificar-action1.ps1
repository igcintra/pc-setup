# ============================================================================
# verificar-action1.ps1 — IG Networks / TI
# ----------------------------------------------------------------------------
# PARTE A (etapas 1-5): procura vestigios do agente de acesso remoto Action1,
#   distribuido em abril/2026 disfarcado de "atualizacao da Adobe".
# PARTE B (etapa 6): procura QUALQUER OUTRO programa de acesso remoto/RMM na
#   maquina e CLASSIFICA (esperado x suspeito x observacao) — assim uma passada
#   so responde as duas perguntas: "foi pega na leva de abril?" e "tem algum
#   acesso remoto estranho aqui?".
#
# PROCEDENCIA (18/08/2026): tudo que as etapas 4, 5 e 6 acharem passa por
#   Get-AuthenticodeSignature — responde "quem assinou esse binario?". Binario SEM
#   assinatura ou com HashMismatch (alterado depois de assinado) rodando de pasta de
#   usuario/ProgramData e' sinal forte, nao depende de conhecer o nome da ameaca.
#
# NAO REMOVE NADA e NAO ALTERA CONFIGURACAO. E' so diagnostico — se achar algo,
# a maquina precisa ser preservada como prova.
#
# Nao exige administrador.
#
# Uso:  irm https://raw.githubusercontent.com/igcintra/pc-setup/master/verificar-action1.ps1 | iex
# ============================================================================

$ErrorActionPreference = 'SilentlyContinue'
$CONTADOR = "https://script.google.com/macros/s/AKfycbwZwJrHL2SnECPzx5inz2K5_AVxbVvukXMra0grAgSbVuNjbxeNnP8sLDGdy-Sf2yfvoA/exec"

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  VERIFICACAO DE SEGURANCA - IG Networks" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Isso leva menos de um minuto e nao altera nada no computador." -ForegroundColor Gray
Write-Host ""

$achados = @()   # marcadores do Action1 (Parte A) -> definem ACHADO/LIMPO
$suspeitos = @() # outros acessos remotos que NAO sao padrao da casa (Parte B)
$obs     = @()   # observacoes informativas (esperado/atencao)
$det     = @()
$guid    = ""

function Etapa($n, $texto) { Write-Host ("[{0}/6] {1}" -f $n, $texto) -ForegroundColor White }

# ---- procedencia: quem assinou o binario ----
# Valid        = assinado e cadeia confiavel
# HashMismatch = assinado MAS ALTERADO depois -> nunca e' normal
# NotSigned    = sem assinatura (aceitavel em ferramenta pequena, ruim em servico)
function Procedencia($caminho) {
    if (-not $caminho) { return $null }
    $p = ([string]$caminho).Trim()
    # tarefa agendada guarda o caminho com variavel: %ProgramData%\... -> expandir
    $p = [Environment]::ExpandEnvironmentVariables($p)
    # PathName de servico vem com argumentos: "C:\x\y.exe" -k algo
    if ($p -match '^"([^"]+)"')      { $p = $matches[1] }
    elseif ($p -match '^(.+?\.exe)') { $p = $matches[1] }
    # alvo que e' SCRIPT: .cmd/.bat/.vbs/.js nao carregam assinatura Authenticode.
    # Script agendado rodando de ProgramData/perfil de usuario e' persistencia -> tratar como ruim.
    if ($p -match '\.(cmd|bat|vbs|js|wsf)$') {
        return [pscustomobject]@{ Caminho = $p; Status = 'ScriptSemAssinatura'; Assinante = '' }
    }
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    $s = Get-AuthenticodeSignature -LiteralPath $p
    if (-not $s) { return $null }
    $quem = ''
    if ($s.SignerCertificate) {
        $cn = ($s.SignerCertificate.Subject -split ',' | Where-Object { $_ -match 'CN=' } | Select-Object -First 1)
        $quem = ($cn -replace '.*CN=','').Trim()
    }
    return [pscustomobject]@{ Caminho = $p; Status = [string]$s.Status; Assinante = $quem }
}
function TextoProc($pr) {
    if (-not $pr)                          { return 'nao foi possivel conferir a assinatura' }
    if ($pr.Status -eq 'Valid')            { return "assinado por: $($pr.Assinante)" }
    if ($pr.Status -eq 'HashMismatch')     { return "ALTERADO DEPOIS DE ASSINADO (HashMismatch) - dizia ser: $($pr.Assinante)" }
    if ($pr.Status -eq 'NotSigned')        { return 'SEM ASSINATURA' }
    if ($pr.Status -eq 'ScriptSemAssinatura') { return 'SCRIPT (.cmd/.bat/.vbs) - nao existe assinatura possivel' }
    return "assinatura com problema ($($pr.Status)) - dizia ser: $($pr.Assinante)"
}
# marcador de risco: binario fora do lugar E sem assinatura confiavel
function ProcRuim($pr) {
    if (-not $pr) { return $false }
    return ($pr.Status -ne 'Valid')
}

# ---- [1] chave de registro do agente ----
Etapa 1 "Procurando registro do agente..."
$chaves = @('HKLM:\SOFTWARE\WOW6432Node\Action1','HKLM:\SOFTWARE\Action1')
foreach ($k in $chaves) {
    if (Test-Path $k) {
        $p = Get-ItemProperty $k
        if ($p.'agent.guid') { $guid = $p.'agent.guid' }
        $achados += "registro"
        $det += "[ACHADO] Chave de registro: $k"
        $det += "         agent.guid  = $($p.'agent.guid')"
        $det += "         system.id   = $($p.'system.id')"
    }
}

# ---- [2] entrada de desinstalacao ----
Etapa 2 "Procurando nos programas instalados..."
$raizes = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
)
$instalados = @(Get-ItemProperty $raizes | Where-Object { $_.DisplayName })
$instalados | Where-Object { $_.DisplayName -match 'Action1' } | ForEach-Object {
    $achados += "programa"
    $det += "[ACHADO] Programa instalado: $($_.DisplayName) $($_.DisplayVersion)"
    $det += "         InstallSource   = $($_.InstallSource)"
    $det += "         InstallDate     = $($_.InstallDate)  (atencao: vem do pacote, pode ser falsa)"
    $det += "         SystemComponent = $($_.SystemComponent)"
}

# ---- [3] log do instalador do Windows (a prova mais confiavel) ----
Etapa 3 "Lendo o historico de instalacoes do Windows..."
$ev = Get-WinEvent -FilterHashtable @{LogName='Application'; ProviderName='MsiInstaller'}
if ($ev) {
    $hit = $ev | Where-Object { $_.Message -match 'Action1|400F3B96|Adobe_Update' }
    if ($hit) {
        $achados += "log-instalador"
        $det += "[ACHADO] Log do Windows Installer registra Action1:"
        foreach ($h in ($hit | Sort-Object TimeCreated)) {
            $msg = ($h.Message -replace "`r?`n", ' ' -replace '\s+', ' ')
            if ($msg.Length -gt 200) { $msg = $msg.Substring(0,200) }
            $det += ("         {0}  id={1}  {2}" -f $h.TimeCreated, $h.Id, $msg)
        }
    }
    $maisAntigo = ($ev | Select-Object -Last 1).TimeCreated
    $det += "         (log cobre de $maisAntigo a $(($ev | Select-Object -First 1).TimeCreated), $($ev.Count) eventos)"
    # a leva do incidente foi 09/04/2026: se o log comeca depois disso, a etapa 3 e' cega
    if ($maisAntigo -gt (Get-Date '2026-04-09')) {
        $obs += "log-nao-cobre-abril"
        $det += "         [atencao] o log comeca DEPOIS de 09/04/2026 -> esta etapa nao consegue"
        $det += "                   provar o passado. 'LIMPO' aqui significa 'nada AGORA', nao 'nunca teve'."
    }
} else {
    $det += "[aviso] Nao foi possivel ler o log do Windows Installer."
}

# ---- [4] servico e pastas (esta ativo AGORA?) ----
Etapa 4 "Verificando se esta em execucao..."
$servicos = @(Get-CimInstance Win32_Service)
$svc = $servicos | Where-Object { $_.Name -match 'action1' -or $_.PathName -match 'Action1' }
if ($svc) {
    $achados += "servico-ativo"
    foreach ($s in $svc) {
        $det += "[ACHADO] Servico: $($s.Name) estado=$($s.State) caminho=$($s.PathName)"
        $det += "         procedencia: $(TextoProc (Procedencia $s.PathName))"
    }
}
foreach ($d in @("$env:ProgramFiles\Action1", "${env:ProgramFiles(x86)}\Action1", "$env:ProgramData\Action1")) {
    if (Test-Path $d) {
        $achados += "pasta"
        $it = Get-ChildItem $d -Recurse -File | Sort-Object CreationTime | Select-Object -First 1
        $det += "[ACHADO] Pasta existe: $d  (arquivo mais antigo: $($it.CreationTime))"
    }
}

# ---- [5] instalador na pasta Downloads ----
Etapa 5 "Procurando o instalador na pasta Downloads..."
$dl = Join-Path $env:USERPROFILE 'Downloads'
Get-ChildItem $dl -File -Filter *.msi | Where-Object { $_.Name -match 'action1|adobe|acrobat|reader|update' } | ForEach-Object {
    $achados += "instalador"
    $h = (Get-FileHash $_.FullName -Algorithm SHA256).Hash
    $det += "[ACHADO] Instalador em Downloads: $($_.Name)"
    $det += "         tamanho=$($_.Length)  data=$($_.LastWriteTime)"
    $det += "         SHA256=$h"
    $det += "         procedencia: $(TextoProc (Procedencia $_.FullName))"
}

# ============================================================================
# [6] OUTROS PROGRAMAS DE ACESSO REMOTO / RMM
# ----------------------------------------------------------------------------
# Classe ESPERADO = ferramenta oficial da casa (presenca e' normal).
# Classe SUSPEITO = acesso remoto que NAO faz parte do padrao IGN.
# ============================================================================
Etapa 6 "Procurando outros programas de acesso remoto..."
# a Parte B e' extra: se algo falhar aqui, o veredito do Action1 (Parte A) tem de sair mesmo assim
try {

    $catalogo = @(
        @{ Nome='AnyDesk';                Rx='anydesk';                            Classe='ESPERADO' },
        @{ Nome='ScreenConnect/ConnectWise'; Rx='screenconnect|connectwise';       Classe='SUSPEITO' },
        @{ Nome='Atera';                  Rx='ateraagent|\batera\b';               Classe='SUSPEITO' },
        @{ Nome='Splashtop';              Rx='splashtop';                          Classe='SUSPEITO' },
        @{ Nome='NinjaRMM/NinjaOne';      Rx='ninjarmm|ninjaone';                  Classe='SUSPEITO' },
        @{ Nome='Syncro';                 Rx='syncro(agent|mps)';                  Classe='SUSPEITO' },
        @{ Nome='Level.io';               Rx='level\.io|level-agent';              Classe='SUSPEITO' },
        @{ Nome='PDQ Connect';            Rx='pdq ?connect';                       Classe='SUSPEITO' },
        @{ Nome='Kaseya/VSA';             Rx='kaseya|agentmon';                    Classe='SUSPEITO' },
        @{ Nome='Tactical RMM';           Rx='tacticalrmm|tacticalagent';          Classe='SUSPEITO' },
        @{ Nome='RustDesk';               Rx='rustdesk';                           Classe='SUSPEITO' },
        @{ Nome='MeshCentral/MeshAgent';  Rx='meshagent|meshcentral';              Classe='SUSPEITO' },
        @{ Nome='DWAgent';                Rx='dwagent';                            Classe='SUSPEITO' },
        @{ Nome='UltraViewer';            Rx='ultraviewer';                        Classe='SUSPEITO' },
        @{ Nome='Ammyy Admin';            Rx='ammyy';                              Classe='SUSPEITO' },
        @{ Nome='Supremo';                Rx='supremo';                            Classe='SUSPEITO' },
        @{ Nome='TeamViewer';             Rx='teamviewer';                         Classe='SUSPEITO' },
        @{ Nome='LogMeIn';                Rx='logmein';                            Classe='SUSPEITO' },
        @{ Nome='Zoho Assist';            Rx='zoho ?assist|zaservice';             Classe='SUSPEITO' },
        @{ Nome='GoTo Resolve/Rescue';    Rx='goto ?(resolve|assist)|logmeinrescue'; Classe='SUSPEITO' },
        @{ Nome='Remote Utilities';       Rx='remote ?utilities|rutserv';          Classe='SUSPEITO' },
        @{ Nome='VNC (Real/Ultra/Tight)'; Rx='realvnc|ultravnc|tightvnc|vncserver'; Classe='SUSPEITO' },
        @{ Nome='ngrok / tunel';          Rx='\bngrok\b|localtonet|playit\.gg';    Classe='SUSPEITO' }
    )

    # inventario coletado uma vez so
    $pastas = @()
    foreach ($base in @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramData, $env:LOCALAPPDATA, $env:APPDATA)) {
        if ($base -and (Test-Path $base)) {
            $pastas += Get-ChildItem $base -Directory | Select-Object @{n='N';e={$_.Name}}, @{n='P';e={$_.FullName}}
        }
    }
    $procs = @(Get-Process | Where-Object { $_.Path } | Select-Object Name, Path)

    foreach ($item in $catalogo) {
        $onde = @()
        $bins = @()
        $instalados | Where-Object { $_.DisplayName -match $item.Rx } | ForEach-Object { $onde += "programa instalado: $($_.DisplayName) $($_.DisplayVersion)" }
        $servicos   | Where-Object { $_.Name -match $item.Rx -or $_.DisplayName -match $item.Rx -or $_.PathName -match $item.Rx } | ForEach-Object { $onde += "servico: $($_.Name) [$($_.State)]"; $bins += $_.PathName }
        $pastas     | Where-Object { $_.N -match $item.Rx } | ForEach-Object { $onde += "pasta: $($_.P)" }
        $procs      | Where-Object { $_.Name -match $item.Rx -or $_.Path -match $item.Rx } | ForEach-Object { $onde += "EM EXECUCAO AGORA: $($_.Path)"; $bins += $_.Path }

        if ($onde.Count) {
            $onde = $onde | Select-Object -Unique
            if ($item.Classe -eq 'ESPERADO') {
                $obs += "esperado:$($item.Nome)"
                $det += "[esperado] $($item.Nome) — ferramenta oficial da casa"
            } else {
                $suspeitos += "rmm:$($item.Nome)"
                $det += "[SUSPEITO] $($item.Nome) — acesso remoto fora do padrao IGN"
            }
            foreach ($o in $onde) { $det += "           $o" }
            foreach ($b in ($bins | Where-Object { $_ } | Select-Object -Unique)) {
                $pr = Procedencia $b
                $det += "           procedencia: $(TextoProc $pr)"
                if ($pr -and $pr.Status -eq 'HashMismatch') {
                    $suspeitos += "binario-alterado:$($item.Nome)"
                    $det += "           [SUSPEITO] binario alterado depois de assinado"
                }
            }
        }
    }

    # --- heuristica: servico ou tarefa rodando de ProgramData / pasta do usuario ---
    # software honesto quase nunca mora ai. Whitelist do que e' normal na frota.
    # Defender/MpDefender/NisSrv REALMENTE moram em ProgramData -> sem isso a heuristica
    # apitava em 18/18 maquinas (medido no 1o lote de 18 relatorios, 18/08/2026).
    $normal = 'dropbox|onedrive|slack|teams|zoom|chrome|edge|adobe|docker|steam|spotify|discord|cursor|postman|whatsapp|google|anydesk|keepass|expressvpn|nvidia|intel|lenovo|dell|edgeupdate|windows defender|mpdefendercoreservice|msmpeng|nissrv|mdcoresvc|myasus'
    $servicos | Where-Object { ($_.PathName -match 'ProgramData|\\Users\\') -and ($_.PathName -notmatch $normal) } | ForEach-Object {
        $det += "[atencao] Servico rodando de local incomum: $($_.Name) [$($_.State)]"
        $det += "           $($_.PathName)"
        $pr = Procedencia $_.PathName
        $det += "           procedencia: $(TextoProc $pr)"
        if (ProcRuim $pr) {
            $suspeitos += "servico-fora-do-lugar-sem-assinatura"
            $det += "           [SUSPEITO] fora do lugar E sem assinatura confiavel"
        } else {
            $obs += "servico-fora-do-lugar"
        }
    }
    Get-ScheduledTask | Where-Object { $_.State -ne 'Disabled' } | ForEach-Object {
        $exec = ($_.Actions | Where-Object { $_.Execute }).Execute -join ' '
        if ($exec -match 'ProgramData|\\Users\\' -and $exec -notmatch $normal) {
            $det += "[atencao] Tarefa agendada de local incomum: $($_.TaskName)"
            $det += "           $exec"
            $pr = Procedencia $exec
            $det += "           procedencia: $(TextoProc $pr)"
            if (ProcRuim $pr) {
                $suspeitos += "tarefa-fora-do-lugar-sem-assinatura"
                $det += "           [SUSPEITO] fora do lugar E sem assinatura confiavel"
            } else {
                $obs += "tarefa-fora-do-lugar"
            }
        }
    }

    # --- AnyDesk e' nosso, mas senha fixa deixa entrar SEM ninguem aceitar ---
    $adConf = @("$env:ProgramData\AnyDesk\service.conf", "$env:APPDATA\AnyDesk\service.conf")
    foreach ($c in $adConf) {
        if (Test-Path $c) {
            $conf = Get-Content $c
            if ($conf -match 'pwd_hash|pwd_salt') {
                $obs += "anydesk-senha-fixa"
                $det += "[atencao] AnyDesk com senha de acesso NAO VIGIADO configurada ($c)."
                $det += "           Permite entrar sem a pessoa aceitar. Conferir se foi a TI que definiu."
            }
        }
    }

    # --- acesso remoto nativo do Windows (RDP) ---
    $rdp = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections).fDenyTSConnections
    if ($rdp -eq 0) {
        $obs += "rdp-ligado"
        $det += "[atencao] Area de Trabalho Remota (RDP) esta HABILITADA nesta maquina."
    }

} catch {
    $det += "[aviso] A varredura de outros acessos remotos falhou nesta maquina: $($_.Exception.Message)"
    $det += "        (o resultado do Action1 acima continua valido)"
}

# ---- veredito ----
$achados   = $achados   | Select-Object -Unique
$suspeitos = $suspeitos | Select-Object -Unique
$obs       = $obs       | Select-Object -Unique

if     ($achados.Count -gt 0)   { $veredito = "ACHADO" }
elseif ($suspeitos.Count -gt 0) { $veredito = "REVISAR" }
else                            { $veredito = "LIMPO" }

$marcadores = @($achados) + @($suspeitos) + @($obs) | Where-Object { $_ }

$desktop = [Environment]::GetFolderPath("Desktop")
if (-not $desktop -or -not (Test-Path $desktop)) { $desktop = "$env:USERPROFILE\Desktop" }
# nome com USUARIO + PC + data: quando varias pessoas mandarem, da p/ saber de quem e' cada um
$limpaU = ($env:USERNAME  -replace '[^A-Za-z0-9._-]','')
$limpaH = ($env:COMPUTERNAME -replace '[^A-Za-z0-9._-]','')
$nomeArq = "verificacao-seguranca-$limpaU-$limpaH-$(Get-Date -Format 'yyyy-MM-dd').txt"
$arq     = Join-Path $desktop $nomeArq
$cab = @(
    "VERIFICACAO DE SEGURANCA - IG Networks",
    "Data......: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')",
    "Computador: $env:COMPUTERNAME",
    "Usuario...: $env:USERNAME",
    "Windows...: $((Get-CimInstance Win32_OperatingSystem).Caption) build $((Get-CimInstance Win32_OperatingSystem).BuildNumber)",
    "VEREDITO..: $veredito",
    "  ACHADO  = vestigio do agente do incidente de abril (Action1)",
    "  REVISAR = sem Action1, mas ha outro acesso remoto fora do padrao",
    "  LIMPO   = nada dos dois",
    "Marcadores: $($marcadores -join ', ')",
    "",
    "----- DETALHES -----"
)
try {
    ($cab + $det) | Set-Content -Path $arq -Encoding UTF8 -ErrorAction Stop
} catch {
    # Desktop bloqueado/OneDrive fora do ar: nao perder o relatorio
    $arq = Join-Path $env:USERPROFILE $nomeArq
    try { ($cab + $det) | Set-Content -Path $arq -Encoding UTF8 -ErrorAction Stop }
    catch { $arq = Join-Path $env:TEMP $nomeArq; ($cab + $det) | Set-Content -Path $arq -Encoding UTF8 }
}

Write-Host ""
if ($veredito -eq "ACHADO") {
    Write-Host "=========================================" -ForegroundColor Red
    Write-Host "  ENCONTRAMOS ALGO - AVISE A TI" -ForegroundColor Red
    Write-Host "=========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "NAO desinstale e NAO apague nada." -ForegroundColor Yellow
    Write-Host "Mande para a TI o arquivo que ficou no seu Desktop:" -ForegroundColor Yellow
    Write-Host "  $nomeArq" -ForegroundColor White
    Write-Host ""
    Write-Host "Marcadores encontrados: $($achados -join ', ')" -ForegroundColor Gray
} elseif ($veredito -eq "REVISAR") {
    Write-Host "=========================================" -ForegroundColor Yellow
    Write-Host "  NADA DO INCIDENTE - MAS A TI VAI CONFERIR" -ForegroundColor Yellow
    Write-Host "=========================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "O programa do incidente NAO esta nesta maquina." -ForegroundColor Gray
    Write-Host "Achamos outro programa de acesso remoto que pode ser normal" -ForegroundColor Gray
    Write-Host "(alguem da TI pode ter instalado) - so precisa ser conferido." -ForegroundColor Gray
    Write-Host ""
    Write-Host "Mande para a TI o arquivo que ficou no seu Desktop:" -ForegroundColor Yellow
    Write-Host "  $nomeArq" -ForegroundColor White
    Write-Host "Nao precisa desinstalar nada." -ForegroundColor Gray
} else {
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host "  ESTA MAQUINA ESTA LIMPA" -ForegroundColor Green
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Nada a fazer. Obrigado pela colaboracao!" -ForegroundColor Gray
    if ($obs.Count) { Write-Host "(o relatorio traz algumas observacoes de rotina para a TI)" -ForegroundColor DarkGray }
}
Write-Host ""
Write-Host "Relatorio salvo em: $arq" -ForegroundColor DarkGray

# ---- envia o resultado para a planilha central ----
function Enc($s) { if ($null -eq $s) { return "" } ; return [uri]::EscapeDataString([string]$s) }
$url = "$CONTADOR" + "?script=verificar-action1" +
       "&host="     + (Enc $env:COMPUTERNAME) +
       "&usuario="  + (Enc $env:USERNAME) +
       "&veredito=" + (Enc $veredito) +
       "&marcadores=" + (Enc ($marcadores -join '|')) +
       "&guid="     + (Enc $guid)
try { Invoke-RestMethod -Uri $url -TimeoutSec 20 | Out-Null } catch { }
Write-Host ""
