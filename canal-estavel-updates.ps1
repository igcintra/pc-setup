# ============================================================
#  canal-estavel-updates.ps1
#  Tira o PC do canal ANTECIPADO de atualizacoes do Windows 11
#  ("Obtenha as atualizacoes mais recentes assim que estiverem disponiveis").
#
#  POR QUE: em 3 maquinas diagnosticadas (30/07 a 03/08/2026) o mesmo estrago
#  apareceu — build 26200 com componente corrompido da build anterior
#  (userexperience-oobe 26100.1591): DISM acusando "repositorio reparavel",
#  updates falhando e horas de reparo (UTI). Maquina de trabalho deve ficar
#  no canal normal, que recebe o mesmo update algumas semanas depois, testado.
#
#  Precisa de ADMIN (chave de maquina). Nao desinstala nada nem reverte builds:
#  o PC para de ANTECIPAR as proximas. Reversivel (basta ligar o botao na UI).
#
#  Uso:
#    irm https://raw.githubusercontent.com/igcintra/pc-setup/master/canal-estavel-updates.ps1 | iex
# ============================================================

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Elevando para Administrador..." -ForegroundColor Yellow
    Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-Command',
        "irm https://raw.githubusercontent.com/igcintra/pc-setup/master/canal-estavel-updates.ps1 | iex"
    return
}

Write-Host "`n=== Canal de atualizacoes: ANTECIPADO -> ESTAVEL ===" -ForegroundColor Cyan

$reg = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"
if (-not (Test-Path $reg)) { New-Item -Path $reg -Force | Out-Null }

$antes = (Get-ItemProperty -Path $reg -Name "IsContinuousInnovationOptedIn" -ErrorAction SilentlyContinue).IsContinuousInnovationOptedIn
Write-Host ("  estado atual: " + $(if ($antes -eq 1) { "ANTECIPADO (vamos desligar)" } else { "ja estava no canal estavel" }))

Set-ItemProperty -Path $reg -Name "IsContinuousInnovationOptedIn" -Value 0 -Type DWord
$depois = (Get-ItemProperty -Path $reg -Name "IsContinuousInnovationOptedIn").IsContinuousInnovationOptedIn

if ($depois -eq 0) {
    Write-Host "  [ok] updates antecipados DESATIVADOS" -ForegroundColor Green
} else {
    Write-Host "  [!] nao aplicou — conferir manualmente" -ForegroundColor Red
}

# build atual, so p/ registro no atendimento
$b = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion")
Write-Host ("`n  Windows atual: " + $b.ProductName + " build " + $b.CurrentBuildNumber + "." + $b.UBR)
Write-Host "  (a build atual NAO volta atras — o PC so para de antecipar as proximas)" -ForegroundColor Yellow

Write-Host "`nConfira em: Configuracoes > Windows Update — a chave 'Obtenha as atualizacoes" -ForegroundColor White
Write-Host "mais recentes assim que estiverem disponiveis' deve aparecer DESATIVADA." -ForegroundColor White
Write-Host "=== PRONTO ===" -ForegroundColor Green
