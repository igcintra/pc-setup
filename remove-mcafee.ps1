# ============================================================
#  remove-mcafee.ps1 - arranca o McAfee de vez (inclusive LiveSafe /
#  Total Protection, que o winget sozinho NAO tira).
#  Roda direto da web, se eleva sozinho:
#    irm https://raw.githubusercontent.com/igcintra/pc-setup/master/remove-mcafee.ps1 | iex
#  Para PC ja entregue, que saiu da fabrica com McAfee e continua notificando.
#  Depois de rodar: REINICIAR. Parte do McAfee so morre no reboot.
# ============================================================

$URL = 'https://raw.githubusercontent.com/igcintra/pc-setup/master/remove-mcafee.ps1'

# ---- auto-elevacao (re-baixa a si mesmo elevado, sem arquivo) ----
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host 'Elevando para Administrador...' -ForegroundColor Yellow
    Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-Command',"irm $URL | iex"
    return
}

$Host.UI.RawUI.WindowTitle = 'Removendo McAfee - nao feche esta janela'
$logDir = Join-Path ([Environment]::GetFolderPath('Desktop')) 'UTI-backup'
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
Start-Transcript -Path (Join-Path $logDir 'remove-mcafee-log.txt') -Append | Out-Null
Write-Host ""
Write-Host "===== remove-mcafee - $(Get-Date) =====" -ForegroundColor Cyan
Write-Host "PC: $env:COMPUTERNAME" -ForegroundColor DarkGray

# nao deixar o PC dormir no meio (o MCPR leva minutos)
$sig = '[DllImport("kernel32.dll")] public static extern uint SetThreadExecutionState(uint f);'
try { (Add-Type -MemberDefinition $sig -Name S -Namespace W -PassThru)::SetThreadExecutionState(0x80000003) | Out-Null } catch {}

# ============================================================
#  McAfee: exterminio completo  (rev. 2026-08-13)
#  Antes este bloco dava conta do McAfee OEM "leve" (WebAdvisor / WPS),
#  mas NAO do LiveSafe / Total Protection (antivirus completo). 4 furos:
#   1) timeout de 15s matava o winget no meio (LiveSafe leva minutos)
#   2) exit code do winget nunca era lido -> imprimia "OK" sem remover nada
#   3) UninstallString do LiveSafe e' GUI (mcuihost.exe): ignora /quiet e,
#      com -WindowStyle Hidden, vira assistente invisivel esperando clique
#   4) self-protection (mfehidk/mfevtp) barra Stop-Service e Remove-Item
#  A peca que faltava: MCPR (mccleanup.exe), removedor oficial da McAfee.
#  No fim o bloco VERIFICA e diz a verdade sobre o que sobrou.
# ============================================================
$mcRegKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
)
$mcPastas = @("$env:ProgramFiles\McAfee", "${env:ProgramFiles(x86)}\McAfee",
              "$env:ProgramData\McAfee", "$env:LOCALAPPDATA\McAfee", "$env:APPDATA\McAfee")

function Get-McAfeeInstalado {
    [pscustomobject]@{
        Programas = @(Get-ItemProperty $mcRegKeys -ErrorAction SilentlyContinue |
                      Where-Object { $_.DisplayName -like "*McAfee*" })
        Servicos  = @(Get-CimInstance Win32_Service -ErrorAction SilentlyContinue |
                      Where-Object { $_.DisplayName -like "*McAfee*" -or $_.Name -like "mfe*" -or $_.PathName -like "*\McAfee\*" })
        Pastas    = @($mcPastas | Where-Object { Test-Path $_ })
        Tarefas   = @(Get-ScheduledTask -ErrorAction SilentlyContinue |
                      Where-Object { $_.TaskName -like "*McAfee*" -or $_.TaskPath -like "*McAfee*" })
        Appx      = @(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue |
                      Where-Object { $_.Name -like "*McAfee*" })
    }
}

function Kill-McAfee {
    Get-Process -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -match "mcafee|mcshield|mcuicnt|modulecore|mmsshost|mcpvtray|webadvisor|mcinstaller|mfemms|mfevtps?|mcods|mfefire|mfetp|protectedmodulehost|mccspservicehost|mcapexe|mfeann|mcpltsvc|msksrvr|mcuihost"
    } | Stop-Process -Force -ErrorAction SilentlyContinue
}

$mc0 = Get-McAfeeInstalado
if (-not ($mc0.Programas -or $mc0.Servicos -or $mc0.Pastas -or $mc0.Tarefas -or $mc0.Appx)) {
    Write-Host "  McAfee: nada instalado neste PC" -ForegroundColor Green
} else {
    Write-Host "  McAfee detectado:" -ForegroundColor Yellow
    foreach ($pg in ($mc0.Programas | Select-Object -ExpandProperty DisplayName -Unique)) {
        Write-Host "    - $pg" -ForegroundColor DarkGray
    }
    if ($mc0.Servicos) { Write-Host "    - $($mc0.Servicos.Count) servico(s), $($mc0.Tarefas.Count) tarefa(s) agendada(s)" -ForegroundColor DarkGray }

    # -- GUARDA: nao deixar o MCPR arrancar antivirus CORPORATIVO (ePO/ENS/Trellix) --
    # MCPR remove produto de CONSUMIDOR. Se a maquina for gerenciada, so avisa.
    $mcCorp = @($mc0.Servicos | Where-Object {
        $_.Name -in @("masvc","macmnsvc","McAfeeFramework") -or $_.DisplayName -match "Trellix|Endpoint Security|ePolicy"
    })
    if ($mcCorp) {
        Write-Host "  !! McAfee/Trellix CORPORATIVO detectado ($($mcCorp[0].DisplayName))." -ForegroundColor Red
        Write-Host "     Pulando o MCPR: ele arrancaria o antivirus gerenciado. Tratar com o time de seguranca." -ForegroundColor Red
    }

    # -- 1. derrubar servicos e processos (best effort: self-protection pode barrar) --
    Write-Host "  McAfee: parando servicos..." -ForegroundColor Yellow
    foreach ($s in $mc0.Servicos) {
        Stop-Service -Name $s.Name -Force -ErrorAction SilentlyContinue
        Set-Service -Name $s.Name -StartupType Disabled -ErrorAction SilentlyContinue
    }
    Kill-McAfee

    # -- 2. tarefas agendadas (raiz das notificacoes: Anti-tracker, Health Check...) --
    $tasksRemoved = 0
    foreach ($t in $mc0.Tarefas) {
        try { Unregister-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -Confirm:$false -ErrorAction Stop; $tasksRemoved++ } catch {}
    }
    if ($tasksRemoved -gt 0) { Write-Host "    $tasksRemoved tarefa(s) agendada(s) apagada(s)" -ForegroundColor DarkGray }

    # -- 3. Appx (OEM novo empacota "McAfee Security"/"Personal Security" da Store) --
    foreach ($ap in $mc0.Appx) {
        Remove-AppxPackage -Package $ap.PackageFullName -AllUsers -ErrorAction SilentlyContinue
    }
    Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like "*McAfee*" } |
        ForEach-Object { Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction SilentlyContinue }
    if ($mc0.Appx) { Write-Host "    $($mc0.Appx.Count) pacote(s) Appx McAfee removido(s)" -ForegroundColor DarkGray }

    # -- 4. winget: timeout de verdade + LEITURA DO EXIT CODE (furos 1 e 2) --
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        foreach ($mid in @("McAfee.WebAdvisor","McAfee.McAfee","McAfee.LiveSafe","McAfee.TotalProtection","McAfee.TrueKey","McAfee.SecurityScan")) {
            Kill-McAfee
            Write-Host "  McAfee: winget $mid..." -ForegroundColor Yellow -NoNewline
            $p = Start-Process "winget" -ArgumentList "uninstall --id $mid -e --silent --force --disable-interactivity --accept-source-agreements" -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
            if (-not $p) { Write-Host " winget nao rodou" -ForegroundColor Gray; continue }
            if (-not $p.WaitForExit(600000)) {          # 10 min: LiveSafe nao sai em 15s
                Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
                Kill-McAfee
                Write-Host " TIMEOUT 10min (abortado)" -ForegroundColor Red
            } elseif ($p.ExitCode -eq 0) {
                Write-Host " removido" -ForegroundColor Green
            } elseif ($p.ExitCode -eq -1978335189 -or $p.ExitCode -eq -1978335212) {
                Write-Host " nao instalado" -ForegroundColor Gray      # 0x8A15002B / 0x8A150014
            } else {
                Write-Host (" falhou (exit {0})" -f $p.ExitCode) -ForegroundColor Yellow
            }
        }
    }

    # -- 5. UninstallString do registro, agora SEM cair na armadilha da GUI (furo 3) --
    foreach ($app in (Get-ItemProperty $mcRegKeys -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*McAfee*" })) {
        $unins = $app.QuietUninstallString; if (-not $unins) { $unins = $app.UninstallString }
        if (-not $unins) { continue }
        if ($unins -match 'MsiExec(\.exe)?\s.*(\{[0-9A-Fa-f\-]{36}\})') {
            # MSI: silencioso de verdade
            $p = Start-Process "msiexec.exe" -ArgumentList "/X$($matches[2]) /qn /norestart" -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
            if ($p) {
                if (-not $p.WaitForExit(600000)) { try { $p.Kill() } catch {} }
                Write-Host "    msi removido: $($app.DisplayName)" -ForegroundColor DarkGray
            }
        } elseif ($app.QuietUninstallString) {
            if ($unins -match '^"([^"]+)"\s*(.*)$') { $exe = $matches[1]; $argStr = $matches[2] }
            else { $tk = $unins -split ' ',2; $exe = $tk[0]; $argStr = if ($tk.Count -gt 1) { $tk[1] } else { '' } }
            if (Test-Path $exe) {
                $p = Start-Process -FilePath $exe -ArgumentList $argStr -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
                if ($p) {
                    if (-not $p.WaitForExit(600000)) { try { $p.Kill() } catch {} }
                    Write-Host "    desinstalado: $($app.DisplayName)" -ForegroundColor DarkGray
                }
            }
        } else {
            # sem QuietUninstallString = assistente grafico. Rodar escondido so trava.
            Write-Host "    $($app.DisplayName): uninstaller e' GUI -> deixado para o MCPR" -ForegroundColor DarkGray
        }
    }

    # -- 6. MCPR / mccleanup.exe: removedor OFICIAL, unico que tira o LiveSafe --
    $mcRestante = Get-McAfeeInstalado
    if (-not $mcCorp -and ($mcRestante.Programas -or $mcRestante.Servicos -or $mcRestante.Pastas)) {
        Write-Host "  McAfee: sobrou produto -> rodando o MCPR oficial (pode levar minutos)..." -ForegroundColor Yellow
        $mcprExe = Join-Path $env:TEMP "MCPR.exe"
        $baixou = $false
        foreach ($u in @("https://download.mcafee.com/molbin/iss-loc/SupportTools/MCPR/MCPR.exe",
                         "https://download.mcafee.com/products/licensed/cust_support_patches/MCPR.exe")) {
            try {
                Remove-Item $mcprExe -Force -ErrorAction SilentlyContinue
                & curl.exe -sL --max-time 300 -o $mcprExe $u 2>$null
                if ((Test-Path $mcprExe) -and ((Get-Item $mcprExe).Length -gt 5MB)) { $baixou = $true; break }
            } catch {}
        }
        if (-not $baixou) {
            Write-Host "    MCPR: download falhou (sem internet?) - baixar de mcafee.com/support e rodar a mao" -ForegroundColor Yellow
        } else {
            # MCPR.exe e' um SFX NSIS: extrai o mccleanup.exe em %TEMP%\ns*.tmp e
            # abre o assistente. O silencio esta no mccleanup, nao no wrapper.
            $sfx = Start-Process $mcprExe -ArgumentList "/S" -PassThru -WindowStyle Minimized -ErrorAction SilentlyContinue
            $clean = $null; $limite = (Get-Date).AddSeconds(120)
            while ((Get-Date) -lt $limite -and -not $clean) {
                Start-Sleep -Seconds 3
                $clean = Get-ChildItem $env:TEMP -Directory -ErrorAction SilentlyContinue |
                         Where-Object { $_.Name -like "ns*.tmp" -or $_.Name -like "*McAfee*" -or $_.Name -like "*MCPR*" } |
                         ForEach-Object { Get-ChildItem $_.FullName -Filter "mccleanup.exe" -Recurse -ErrorAction SilentlyContinue } |
                         Select-Object -First 1
            }
            if (-not $clean) {
                if ($sfx) { try { $sfx | Stop-Process -Force -ErrorAction SilentlyContinue } catch {} }
                Write-Host "    MCPR: mccleanup.exe nao apareceu em 2 min - rodar $mcprExe a mao" -ForegroundColor Yellow
            } else {
                # copiar ANTES de matar o SFX: o NSIS apaga o proprio $PLUGINSDIR ao sair
                $work = Join-Path $env:TEMP "MCPR-work"
                Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
                Copy-Item $clean.Directory.FullName $work -Recurse -Force -ErrorAction SilentlyContinue
                if ($sfx) { try { $sfx | Stop-Process -Force -ErrorAction SilentlyContinue } catch {} }
                Get-Process -Name "MCPR","mccleanup" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                $exeClean = Join-Path $work "mccleanup.exe"
                if (Test-Path $exeClean) {
                    $mods = "StopServices,MFSY,PEF,MXD,CSP,Sustainability,MOCP,MFP,APPSTATS,Auth,EMproxy,FWdiver,HW,MAS,MAT,MBK,MCPR,McProxy,McSvcHost,VUL,MHN,MNA,MOBK,MPFP,MPFPCU,MPS,SHRED,MPSCU,MQC,MQCCU,MSAD,MSHR,MSK,MSKCU,MWL,NMC,RedirSvc,VS,REMEDIATION,MSC,YAP,TRUEKEY,LAM,PCB,Symlink,SafeConnect,MGS,WMIRemover,RESIDUE"
                    $pc = Start-Process $exeClean -ArgumentList "-p","$mods","-v","-s" -PassThru -WindowStyle Hidden -WorkingDirectory $work -ErrorAction SilentlyContinue
                    if ($pc) {
                        if (-not $pc.WaitForExit(900000)) { try { $pc.Kill() } catch {}; Write-Host "    MCPR: timeout de 15 min" -ForegroundColor Yellow }
                        else { Write-Host "    MCPR concluido (exit $($pc.ExitCode))" -ForegroundColor DarkGray }
                    }
                } else {
                    Write-Host "    MCPR: falha ao copiar o mccleanup - rodar $mcprExe a mao" -ForegroundColor Yellow
                }
            }
        }
    }

    # -- 7. faxina final: servicos zumbis, pastas, chaves de inicializacao --
    Kill-McAfee
    foreach ($s in (Get-CimInstance Win32_Service -ErrorAction SilentlyContinue |
                    Where-Object { $_.DisplayName -like "*McAfee*" -or $_.Name -like "mfe*" -or $_.PathName -like "*\McAfee\*" })) {
        if ($s.Name -in @("masvc","macmnsvc","McAfeeFramework")) { continue }   # corporativo: nao tocar
        Stop-Service -Name $s.Name -Force -ErrorAction SilentlyContinue
        Start-Process sc.exe -ArgumentList "delete `"$($s.Name)`"" -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue | Out-Null
    }
    foreach ($p in $mcPastas) { if (Test-Path $p) { Remove-Item $p -Recurse -Force -ErrorAction SilentlyContinue } }
    foreach ($hive in @("HKLM:","HKCU:")) {
        $k = "$hive\Software\Microsoft\Windows\CurrentVersion\Run"
        if (Test-Path $k) {
            (Get-Item $k).Property | Where-Object { $_ -like "*McAfee*" -or $_ -like "*mcui*" } |
                ForEach-Object { Remove-ItemProperty -Path $k -Name $_ -ErrorAction SilentlyContinue }
        }
    }

    # -- 8. VERIFICACAO: nunca mais dizer "OK" sem ter removido (furo 2) --
    $mcFim = Get-McAfeeInstalado
    if ($mcFim.Programas -or $mcFim.Servicos -or $mcFim.Pastas) {
        Write-Host "  !! McAfee AINDA PRESENTE apos a limpeza:" -ForegroundColor Red
        foreach ($pg in ($mcFim.Programas | Select-Object -ExpandProperty DisplayName -Unique)) { Write-Host "     programa: $pg" -ForegroundColor Red }
        foreach ($s in $mcFim.Servicos) { Write-Host "     servico: $($s.Name)" -ForegroundColor Red }
        foreach ($p in $mcFim.Pastas) { Write-Host "     pasta: $p" -ForegroundColor Red }
        Write-Host "     -> REINICIE e rode de novo: parte do McAfee so sai apos o reboot." -ForegroundColor Yellow
    } else {
        Write-Host "  McAfee: removido e verificado (0 programas, 0 servicos, 0 pastas)" -ForegroundColor Green
    }
    try {
        $def = Get-MpComputerStatus -ErrorAction Stop
        Write-Host ("    Defender apos a remocao: tempo real={0} / antivirus={1} (assume o posto no reboot)" -f $def.RealTimeProtectionEnabled, $def.AntivirusEnabled) -ForegroundColor DarkGray
    } catch {}
}
Write-Host "  McAfee: concluido" -ForegroundColor Green

Write-Host ""
Write-Host "REINICIE O PC agora - parte da remocao so se completa no reboot." -ForegroundColor Yellow
Write-Host "Depois do reboot, rode este script de novo: ele so confirma se sobrou algo." -ForegroundColor Yellow
Stop-Transcript | Out-Null
