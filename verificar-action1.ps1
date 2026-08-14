# ============================================================================
# verificar-action1.ps1 — IG Networks / TI
# ----------------------------------------------------------------------------
# Verifica se esta maquina tem vestigios do agente de acesso remoto Action1,
# distribuido em abril/2026 disfarcado de "atualizacao da Adobe".
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

$achados = @()
$det     = @()
$guid    = ""

function Etapa($n, $texto) { Write-Host ("[{0}/5] {1}" -f $n, $texto) -ForegroundColor White }

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
foreach ($r in $raizes) {
    Get-ItemProperty $r | Where-Object { $_.DisplayName -match 'Action1' } | ForEach-Object {
        $achados += "programa"
        $det += "[ACHADO] Programa instalado: $($_.DisplayName) $($_.DisplayVersion)"
        $det += "         InstallSource   = $($_.InstallSource)"
        $det += "         InstallDate     = $($_.InstallDate)  (atencao: vem do pacote, pode ser falsa)"
        $det += "         SystemComponent = $($_.SystemComponent)"
    }
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
    $det += "         (log cobre de $(($ev | Select-Object -Last 1).TimeCreated) a $(($ev | Select-Object -First 1).TimeCreated), $($ev.Count) eventos)"
} else {
    $det += "[aviso] Nao foi possivel ler o log do Windows Installer."
}

# ---- [4] servico e pastas (esta ativo AGORA?) ----
Etapa 4 "Verificando se esta em execucao..."
$svc = Get-CimInstance Win32_Service | Where-Object { $_.Name -match 'action1' -or $_.PathName -match 'Action1' }
if ($svc) {
    $achados += "servico-ativo"
    foreach ($s in $svc) { $det += "[ACHADO] Servico: $($s.Name) estado=$($s.State) caminho=$($s.PathName)" }
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
}

# ---- veredito ----
$achados = $achados | Select-Object -Unique
$sujo    = ($achados.Count -gt 0)
if ($sujo) { $veredito = "ACHADO" } else { $veredito = "LIMPO" }

$desktop = [Environment]::GetFolderPath("Desktop")
$arq     = Join-Path $desktop "verificacao-seguranca.txt"
$cab = @(
    "VERIFICACAO DE SEGURANCA - IG Networks",
    "Data......: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')",
    "Computador: $env:COMPUTERNAME",
    "Usuario...: $env:USERNAME",
    "Windows...: $((Get-CimInstance Win32_OperatingSystem).Caption) build $((Get-CimInstance Win32_OperatingSystem).BuildNumber)",
    "VEREDITO..: $veredito",
    "Marcadores: $($achados -join ', ')",
    "",
    "----- DETALHES -----"
)
($cab + $det) | Set-Content -Path $arq -Encoding UTF8

Write-Host ""
if ($sujo) {
    Write-Host "=========================================" -ForegroundColor Red
    Write-Host "  ENCONTRAMOS ALGO - AVISE A TI" -ForegroundColor Red
    Write-Host "=========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "NAO desinstale e NAO apague nada." -ForegroundColor Yellow
    Write-Host "Mande para a TI o arquivo que ficou no seu Desktop:" -ForegroundColor Yellow
    Write-Host "  verificacao-seguranca.txt" -ForegroundColor White
    Write-Host ""
    Write-Host "Marcadores encontrados: $($achados -join ', ')" -ForegroundColor Gray
} else {
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host "  ESTA MAQUINA ESTA LIMPA" -ForegroundColor Green
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Nada a fazer. Obrigado pela colaboracao!" -ForegroundColor Gray
}
Write-Host ""
Write-Host "Relatorio salvo em: $arq" -ForegroundColor DarkGray

# ---- envia o resultado para a planilha central ----
function Enc($s) { if ($null -eq $s) { return "" } ; return [uri]::EscapeDataString([string]$s) }
$url = "$CONTADOR" + "?script=verificar-action1" +
       "&host="     + (Enc $env:COMPUTERNAME) +
       "&usuario="  + (Enc $env:USERNAME) +
       "&veredito=" + (Enc $veredito) +
       "&marcadores=" + (Enc ($achados -join '|')) +
       "&guid="     + (Enc $guid)
try { Invoke-RestMethod -Uri $url -TimeoutSec 20 | Out-Null } catch { }
Write-Host ""
