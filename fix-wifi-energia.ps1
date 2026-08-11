# ============================================================================
# fix-wifi-energia.ps1 - impede o Windows de desligar a placa Wi-Fi - IG Networks
# ----------------------------------------------------------------------------
# Uso (PowerShell comum - ele se AUTO-ELEVA, so clicar "Sim" no UAC):
#   irm https://raw.githubusercontent.com/igcintra/pc-setup/master/fix-wifi-energia.ps1 | iex
# Faz: (1) desliga "permitir que o computador desligue este dispositivo" na
# placa Wi-Fi (via MSPower_DeviceEnable, casado pelo PnPDeviceID real da NIC);
# (2) MIMO Power Save -> Sem SMPS; (3) mostra o estado final como prova.
# Para o sintoma "conecta e cai / desconectada pelo driver" (Netwtw10 6062).
# Caso de origem: aamorin (Lenovo, Intel AC 9560, Win11 24H2), 11/08/2026.
# ============================================================================
$ErrorActionPreference = 'Continue'
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Elevando... clique SIM na janela que vai abrir." -ForegroundColor Yellow
    Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-Command',"irm https://raw.githubusercontent.com/igcintra/pc-setup/master/fix-wifi-energia.ps1 | iex"
    return
}

$wifi = Get-NetAdapter | Where-Object { $_.InterfaceDescription -match 'Wireless|Wi-?Fi|WLAN' } | Select-Object -First 1
if (-not $wifi) { Write-Host "ERRO: nao achei adaptador Wi-Fi." -ForegroundColor Red; Read-Host "Enter para fechar"; return }
Write-Host "Placa: $($wifi.InterfaceDescription)" -ForegroundColor Cyan

# id PnP da placa (ex: PCI\VEN_8086&DEV_A0F0...) -> casa com o InstanceName do MSPower
$pnp = (Get-PnpDevice -FriendlyName $wifi.InterfaceDescription -ErrorAction SilentlyContinue | Select-Object -First 1).InstanceId
$alvo = Get-CimInstance -Namespace root\wmi -ClassName MSPower_DeviceEnable -ErrorAction SilentlyContinue |
        Where-Object { $pnp -and $_.InstanceName -like "$($pnp.Split('\')[0..1] -join '\')*" }
if ($alvo) {
    $alvo | Set-CimInstance -Property @{Enable = $false}
    Write-Host "[1/2] Economia de energia da placa: DESLIGADA" -ForegroundColor Green
} else {
    Write-Host "[1/2] AVISO: nao casei o MSPower da placa (pulei)" -ForegroundColor Yellow
}

Set-NetAdapterAdvancedProperty -Name $wifi.Name -DisplayName 'MIMO - Modo de Economia de Energia' -DisplayValue 'Sem SMPS' -ErrorAction SilentlyContinue
Set-NetAdapterAdvancedProperty -Name $wifi.Name -DisplayName 'MIMO Power Save Mode' -DisplayValue 'No SMPS' -ErrorAction SilentlyContinue
Write-Host "[2/2] MIMO Power Save: Sem SMPS (se a placa tiver a opcao)" -ForegroundColor Green

Write-Host ""; Write-Host "=== PROVA (Enable=False e' o que queremos) ===" -ForegroundColor Cyan
if ($alvo) { Get-CimInstance -Namespace root\wmi -ClassName MSPower_DeviceEnable | Where-Object { $_.InstanceName -eq $alvo.InstanceName } | Select-Object InstanceName, Enable | Format-List }
Get-NetAdapterAdvancedProperty -Name $wifi.Name -ErrorAction SilentlyContinue | Where-Object DisplayName -match 'MIMO' | Select-Object DisplayName, DisplayValue | Format-Table -AutoSize

Write-Host "FEITO! Agora: Windows Update > Atualizacoes opcionais > driver Intel Wireless." -ForegroundColor Green
Read-Host "Enter para fechar"
