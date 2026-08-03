# ============================================================
#  restaurar-notificacoes.ps1  —  desfaz a etapa [8] do setup.ps1
#  (o setup desativa os TOASTS do Windows; isso tambem cala Slack,
#   Teams, Outlook e qualquer app que notifique por toast)
#
#  Caso de origem: Pedro Leao (pleao) — "Slack nao notifica" (03/08/2026)
#
#  Rodar como o PROPRIO usuario (as chaves HKCU sao por perfil!).
#  A parte HKLM precisa de admin — o script avisa se nao tiver.
#
#  Uso (PowerShell):
#    irm https://raw.githubusercontent.com/igcintra/pc-setup/main/restaurar-notificacoes.ps1 | iex
#  ou local:  powershell -ExecutionPolicy Bypass -File .\restaurar-notificacoes.ps1
# ============================================================

Write-Host "`n=== Restaurando notificacoes nativas do Windows ===" -ForegroundColor Cyan

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# --- 1) Toasts do usuario (o principal: e isso que cala o Slack) ---
$reg = "HKCU:\Software\Microsoft\Windows\CurrentVersion\PushNotifications"
if (-not (Test-Path $reg)) { New-Item -Path $reg -Force | Out-Null }
Set-ItemProperty -Path $reg -Name "ToastEnabled" -Value 1 -Type DWord
Write-Host "  [ok] ToastEnabled = 1" -ForegroundColor Green

# --- 2) Politica que bloqueava toast de aplicativo ---
$pol = "HKCU:\Software\Policies\Microsoft\Windows\CurrentVersion\PushNotifications"
if (Test-Path $pol) {
    Remove-ItemProperty -Path $pol -Name "NoToastApplicationNotification" -ErrorAction SilentlyContinue
    Write-Host "  [ok] politica NoToastApplicationNotification removida" -ForegroundColor Green
}

# --- 3) Configuracoes globais de notificacao (toasts + acima da tela de bloqueio) ---
$set = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings"
if (-not (Test-Path $set)) { New-Item -Path $set -Force | Out-Null }
Set-ItemProperty -Path $set -Name "NOC_GLOBAL_SETTING_TOASTS_ENABLED"        -Value 1 -Type DWord
Set-ItemProperty -Path $set -Name "NOC_GLOBAL_SETTING_ALLOW_TOASTS_ABOVE_LOCK" -Value 1 -Type DWord
Set-ItemProperty -Path $set -Name "NOC_GLOBAL_SETTING_ALLOW_NOTIFICATION_SOUND" -Value 1 -Type DWord
Write-Host "  [ok] toasts globais + tela de bloqueio + som" -ForegroundColor Green

# --- 4) Reabilitar o Slack especificamente (se o Windows guardou o app como bloqueado) ---
$apps = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings"
Get-ChildItem $apps -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match 'slack' } | ForEach-Object {
    Set-ItemProperty -Path $_.PSPath -Name "Enabled" -Value 1 -Type DWord -ErrorAction SilentlyContinue
    Write-Host ("  [ok] app reabilitado: " + $_.PSChildName) -ForegroundColor Green
}

# --- 5) Assistente de Foco / Nao Perturbe (se estiver ligado, cala tudo mesmo com toast ON) ---
$focus = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\Windows.FocusAssist"
if (Test-Path $focus) {
    Set-ItemProperty -Path $focus -Name "Enabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
}
Write-Host "  [i] confira tambem: Configuracoes > Sistema > Notificacoes > 'Nao perturbe' DESLIGADO" -ForegroundColor Yellow

# --- 6) Parte de maquina (dicas/consumer features) — precisa admin, e opcional ---
if ($isAdmin) {
    $cc = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
    if (Test-Path $cc) {
        Remove-ItemProperty -Path $cc -Name "DisableSoftLanding" -ErrorAction SilentlyContinue
        Write-Host "  [ok] DisableSoftLanding removido (dicas do Windows voltam)" -ForegroundColor Green
    }
    # DisableWindowsConsumerFeatures fica como esta: e o que segura app promocional/bloatware.
    Write-Host "  [i] DisableWindowsConsumerFeatures MANTIDO de proposito (evita bloatware)" -ForegroundColor Yellow
} else {
    Write-Host "  [i] sem admin: parte HKLM (dicas do Windows) nao alterada — nao afeta o Slack" -ForegroundColor Yellow
}

# --- 7) Reiniciar o Explorer p/ aplicar sem precisar deslogar ---
Write-Host "`nReiniciando o Explorer para aplicar..." -ForegroundColor Cyan
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
if (-not (Get-Process explorer -ErrorAction SilentlyContinue)) { Start-Process explorer }

Write-Host "`n=== PRONTO ===" -ForegroundColor Green
Write-Host "Teste: feche e reabra o Slack, e peca a alguem pra te mandar uma DM." -ForegroundColor White
Write-Host "Se ainda nao notificar, confira no Slack: Preferencias > Notificacoes" -ForegroundColor White
Write-Host "  - 'Notificar sobre' = Todas as mensagens novas (ou Mencoes)" -ForegroundColor White
Write-Host "  - 'Definir agenda de notificacoes' (horario) e 'Pausar notificacoes' DESLIGADOS" -ForegroundColor White
