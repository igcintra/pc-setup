# ============================================================================
# diag-wifi.ps1 - Diagnostico de Wi-Fi que "conecta e cai" - IG Networks
# ----------------------------------------------------------------------------
# Uso (PowerShell comum, nao precisa admin*):
#   irm https://raw.githubusercontent.com/igcintra/pc-setup/master/diag-wifi.ps1 | iex
# Coleta: adaptadores/driver, sinal/banda, economia de energia da placa,
# propriedades avancadas, plano de energia, eventos de queda (8003) e
# erros de rede/driver das ultimas 48h. Salva relatorio no Desktop e ja
# copia pro clipboard - a pessoa so cola no Slack (Ctrl+V).
# *Se a secao de eventos vier vazia, rodar de novo em PowerShell admin.
# Caso de origem: aamorin (Lenovo novo, quedas em 2 redes), 11/08/2026.
# ============================================================================
$r = @()
$r += "=== DIAG WIFI $(Get-Date) - $env:COMPUTERNAME ==="
$r += "--- ADAPTADORES DE REDE ---"
$r += Get-NetAdapter | Select-Object Name,InterfaceDescription,Status,LinkSpeed,DriverVersion,DriverDate,DriverProvider | Format-List | Out-String
$r += "--- INTERFACE WLAN (sinal/canal/banda) ---"
$r += (netsh wlan show interfaces) -join "`r`n"
$r += "--- ECONOMIA DE ENERGIA DO ADAPTADOR (Enable=True: Windows pode DESLIGAR a placa) ---"
$r += Get-CimInstance -Namespace root\wmi -ClassName MSPower_DeviceEnable -ErrorAction SilentlyContinue | Where-Object {$_.InstanceName -match 'PCI'} | Select-Object InstanceName,Enable | Format-List | Out-String
$r += "--- PROPRIEDADES AVANCADAS DA PLACA WI-FI ---"
$r += Get-NetAdapterAdvancedProperty -Name *Wi*Fi*,*WLAN* -ErrorAction SilentlyContinue | Select-Object DisplayName,DisplayValue | Format-Table -AutoSize | Out-String
$r += "--- PLANO DE ENERGIA ---"
$r += (powercfg /getactivescheme) -join "`r`n"
$r += "--- ULTIMAS CONEXOES/QUEDAS DO WI-FI (8001=conectou, 8003=caiu) ---"
$r += Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-WLAN-AutoConfig/Operational';Id=8001,8003} -MaxEvents 25 -ErrorAction SilentlyContinue | Select-Object TimeCreated,Id,@{n='Detalhe';e={($_.Message -split "`r?`n" | Select-String 'SSID|Raz|Reason|Motivo' | Select-Object -First 2) -join ' | '}} | Format-Table -Wrap | Out-String
$r += "--- ERROS DE SISTEMA (rede/driver, ultimas 48h) ---"
$r += Get-WinEvent -FilterHashtable @{LogName='System';Level=1,2,3;StartTime=(Get-Date).AddDays(-2)} -MaxEvents 40 -ErrorAction SilentlyContinue | Where-Object {$_.ProviderName -match 'netw|wlan|ndis|tcpip|wifi'} | Select-Object TimeCreated,ProviderName,Id | Format-Table -AutoSize | Out-String
$out = "$env:USERPROFILE\Desktop\diag-wifi.txt"
$r -join "`r`n" | Out-File $out -Encoding utf8
Get-Content $out -Raw | Set-Clipboard
Write-Host ""
Write-Host "PRONTO! Relatorio salvo em $out e ja COPIADO - e so colar no Slack com Ctrl+V" -ForegroundColor Green
