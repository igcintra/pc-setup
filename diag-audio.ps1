# diag-audio.ps1
# Diagnostico de audio (mic/fone) para Windows 10/11
# Coleta dados em formato markdown e copia pra clipboard.
# Cole o resultado numa IA pedindo: "Esse PC tem problema de microfone, o que voce identifica?"
#
# Uso: irm https://raw.githubusercontent.com/igcintra/pc-setup/main/diag-audio.ps1 | iex
# Repo: github.com/igcintra/pc-setup

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$sb = New-Object System.Text.StringBuilder

function Add-Section {
    param([string]$Title, [string]$Body)
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("## $Title")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine('```')
    [void]$sb.AppendLine($Body)
    [void]$sb.AppendLine('```')
}

# ----------------------------------------------------------------------
# Cabecalho
# ----------------------------------------------------------------------
$os = Get-CimInstance Win32_OperatingSystem
[void]$sb.AppendLine("# Diagnostico de Audio - $env:COMPUTERNAME")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("- Usuario: $env:USERNAME")
[void]$sb.AppendLine("- OS: $($os.Caption) $($os.Version) $($os.OSArchitecture)")
[void]$sb.AppendLine("- Data: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$sb.AppendLine("- Idioma do sistema: $((Get-Culture).Name)")

# ----------------------------------------------------------------------
# 1. Endpoints de audio (PnP)
# ----------------------------------------------------------------------
try {
    $endpoints = Get-PnpDevice -Class AudioEndpoint -ErrorAction Stop |
        Sort-Object Status, FriendlyName |
        Select-Object Status, FriendlyName, Class
    $body = ($endpoints | Format-Table -AutoSize | Out-String).Trim()
} catch { $body = "ERRO: $($_.Exception.Message)" }
Add-Section "Endpoints de Audio (Get-PnpDevice)" $body

# ----------------------------------------------------------------------
# 2. Dispositivos de som (hardware)
# ----------------------------------------------------------------------
try {
    $snd = Get-CimInstance Win32_SoundDevice -ErrorAction Stop |
        Select-Object Name, Status, Manufacturer
    $body = ($snd | Format-Table -AutoSize | Out-String).Trim()
} catch { $body = "ERRO: $($_.Exception.Message)" }
Add-Section "Dispositivos de Som - Hardware" $body

# ----------------------------------------------------------------------
# 3. Drivers de audio (versao e data)
# ----------------------------------------------------------------------
try {
    $drv = Get-CimInstance Win32_PnPSignedDriver -ErrorAction Stop |
        Where-Object { $_.DeviceName -match 'Realtek|Conexant|IDT|Audio|Sound|HD Codec' } |
        Select-Object DeviceName, DriverVersion, @{n='DriverDate';e={
            if ($_.DriverDate) { ($_.DriverDate.Substring(0,8)) } else { '' }
        }}, Manufacturer |
        Sort-Object DeviceName -Unique
    $body = ($drv | Format-Table -AutoSize -Wrap | Out-String).Trim()
} catch { $body = "ERRO: $($_.Exception.Message)" }
Add-Section "Drivers de Audio" $body

# ----------------------------------------------------------------------
# 4. Servicos de audio
# ----------------------------------------------------------------------
try {
    $svcs = Get-Service -Name Audiosrv,AudioEndpointBuilder -ErrorAction Stop |
        Select-Object Name, Status, StartType
    $body = ($svcs | Format-Table -AutoSize | Out-String).Trim()
} catch { $body = "ERRO: $($_.Exception.Message)" }
Add-Section "Servicos de Audio" $body

# ----------------------------------------------------------------------
# 5. Privacidade do microfone (registry)
# ----------------------------------------------------------------------
$privSb = New-Object System.Text.StringBuilder
$paths = @(
    @{Hive='HKCU'; Path='HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone'; Label='User - app store apps'},
    @{Hive='HKLM'; Path='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone'; Label='Machine - app store apps'},
    @{Hive='HKCU'; Path='HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone\NonPackaged'; Label='User - desktop apps'},
    @{Hive='HKLM'; Path='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone\NonPackaged'; Label='Machine - desktop apps'}
)
foreach ($p in $paths) {
    if (Test-Path $p.Path) {
        $v = (Get-ItemProperty -Path $p.Path -ErrorAction SilentlyContinue).Value
        [void]$privSb.AppendLine(("{0,-30} : {1}" -f $p.Label, $v))
    } else {
        [void]$privSb.AppendLine(("{0,-30} : (chave nao existe)" -f $p.Label))
    }
}
Add-Section "Permissoes - Privacidade do Microfone" $privSb.ToString().Trim()

# ----------------------------------------------------------------------
# 6. Painel Realtek (console moderno e classico)
# ----------------------------------------------------------------------
$rtkSb = New-Object System.Text.StringBuilder
$rtkApp = Get-AppxPackage *Realtek* -ErrorAction SilentlyContinue
if ($rtkApp) {
    foreach ($a in $rtkApp) {
        [void]$rtkSb.AppendLine("- Store: $($a.Name) v$($a.Version)")
    }
} else {
    [void]$rtkSb.AppendLine("- Store: nao instalado")
}
$rtkClassic = @(
    'C:\Program Files\Realtek\Audio\HDA\RtkNGUI64.exe',
    'C:\Program Files\Realtek\Audio\HDA\RAVCpl64.exe',
    'C:\Windows\System32\RAVCpl64.exe',
    'C:\Windows\System32\RtkNGUI64.exe'
) | Where-Object { Test-Path $_ }
if ($rtkClassic) {
    foreach ($p in $rtkClassic) { [void]$rtkSb.AppendLine("- Classic: $p") }
} else {
    [void]$rtkSb.AppendLine("- Classic: nao encontrado")
}
Add-Section "Painel Realtek" $rtkSb.ToString().Trim()

# ----------------------------------------------------------------------
# 7. Endpoints com estado real (registry MMDevices)
# Mostra: Active / Disabled / NotPresent / Unplugged
# ----------------------------------------------------------------------
$mmSb = New-Object System.Text.StringBuilder
$stateMap = @{ 1='Active'; 2='Disabled'; 4='NotPresent'; 8='Unplugged' }
foreach ($role in @('Capture','Render')) {
    $base = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$role"
    if (-not (Test-Path $base)) { continue }
    Get-ChildItem $base -ErrorAction SilentlyContinue | ForEach-Object {
        $deviceState = (Get-ItemProperty -Path $_.PSPath -Name DeviceState -ErrorAction SilentlyContinue).DeviceState
        $stateStr = $stateMap[[int]$deviceState]
        if (-not $stateStr) { $stateStr = "Unknown($deviceState)" }
        $propsPath = Join-Path $_.PSPath 'Properties'
        $friendlyName = ''
        $jackSubtype = ''
        if (Test-Path $propsPath) {
            $props = Get-ItemProperty -Path $propsPath -ErrorAction SilentlyContinue
            $friendlyName = $props.'{a45c254e-df1c-4efd-8020-67d146a850e0},2'
            $jackSubtype  = $props.'{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},2'
        }
        [void]$mmSb.AppendLine(("[{0,-7}/{1,-10}] {2}" -f $role, $stateStr, $friendlyName))
    }
}
Add-Section "Endpoints com Estado (registry MMDevices)" $mmSb.ToString().Trim()

# ----------------------------------------------------------------------
# 8. Volume / mute do default capture (best-effort via COM)
# ----------------------------------------------------------------------
$volSb = New-Object System.Text.StringBuilder
try {
    Add-Type -TypeDefinition @"
        using System;
        using System.Runtime.InteropServices;
        [Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IAudioEndpointVolume {
            int RegisterControlChangeNotify(IntPtr p);
            int UnregisterControlChangeNotify(IntPtr p);
            int GetChannelCount(out uint c);
            int SetMasterVolumeLevel(float l, Guid g);
            int SetMasterVolumeLevelScalar(float l, Guid g);
            int GetMasterVolumeLevel(out float l);
            int GetMasterVolumeLevelScalar(out float l);
            int SetChannelVolumeLevel(uint ch, float l, Guid g);
            int SetChannelVolumeLevelScalar(uint ch, float l, Guid g);
            int GetChannelVolumeLevel(uint ch, out float l);
            int GetChannelVolumeLevelScalar(uint ch, out float l);
            int SetMute(bool m, Guid g);
            int GetMute(out bool m);
        }
        [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IMMDevice {
            int Activate(ref Guid id, int clsCtx, IntPtr act, [MarshalAs(UnmanagedType.IUnknown)] out object o);
        }
        [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IMMDeviceEnumerator {
            int NotImpl1();
            int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice ep);
        }
        [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
        public class MMDeviceEnumeratorComObject { }
        public class Helper {
            public static IAudioEndpointVolume GetDefaultCapture() {
                var enumerator = (IMMDeviceEnumerator)(new MMDeviceEnumeratorComObject());
                IMMDevice dev = null;
                enumerator.GetDefaultAudioEndpoint(1, 0, out dev); // 1=eCapture, 0=eConsole
                Guid IID = typeof(IAudioEndpointVolume).GUID;
                object o = null;
                dev.Activate(ref IID, 1, IntPtr.Zero, out o);
                return (IAudioEndpointVolume)o;
            }
        }
"@ -ErrorAction Stop
    $vol = [Helper]::GetDefaultCapture()
    $level = 0.0
    [void]$vol.GetMasterVolumeLevelScalar([ref]$level)
    $mute = $false
    [void]$vol.GetMute([ref]$mute)
    [void]$volSb.AppendLine("Default Capture - Volume: $([math]::Round($level*100,1))%  Mute: $mute")
} catch {
    [void]$volSb.AppendLine("ERRO: $($_.Exception.Message)")
}
Add-Section "Volume / Mute do Microfone Default" $volSb.ToString().Trim()

# ----------------------------------------------------------------------
# Final: copia para clipboard e imprime
# ----------------------------------------------------------------------
$result = $sb.ToString()

try {
    $result | Set-Clipboard
    Write-Host ""
    Write-Host "=========================================================" -ForegroundColor Green
    Write-Host " Diagnostico copiado para a area de transferencia (Ctrl+V)" -ForegroundColor Green
    Write-Host "=========================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Cole numa IA (Claude/ChatGPT) com a pergunta:" -ForegroundColor Yellow
    Write-Host '  "Esse PC tem problema de microfone, o que voce identifica?"' -ForegroundColor Yellow
    Write-Host ""
} catch {
    Write-Host "(nao foi possivel copiar para clipboard - copie manualmente o output abaixo)" -ForegroundColor Yellow
}

Write-Output $result
