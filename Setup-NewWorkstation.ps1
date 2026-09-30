<#
.SYNOPSIS
    Setup skripta za novo delovno postajo - v4 (Optimizirana in izboljsana zanesljivost).

.OPIS
    Windows Update motor je prevzet iz Remont-v3 (COM Windows Update Agent),
    ker je zanesljivejsi od PSWindowsUpdate.

    Dodane optimizacije:
      - GUI obvestila (MessageBox) ob napakah ali preprecitvi ciklanja.
      - Skrajsan timeout za posodobitve (30 min) in avtomatski reset wuauserv storitve ob zataknitvi.
      - Rezervni samodejni zagon preko HKLM RunOnce registra.
      - Onemogocanje Fast Startup (Hitri zagon) za zanesljivo izvajanje ob restartu.

    Koraki:
      1. Predpriprava (power plan, Fast Startup, servisi, WSUS/GPO, prostor)
      2. Obnovitvena tocka "naveza-<datum>"   <- PRED posodobitvami
      3. Windows Update (ciklicno, avtonomno, z restarti)
      4. Gonilniki (OPCIJSKO, -IncludeDrivers)
      5. Aplikacije (winget, z verifikacijo)
      6. Privzeti PDF
      7. DISM RestoreHealth + SFC /scannow

    POPRAVEK: winget install klici zdaj eksplicitno uporabljajo
    "--source winget", ker msstore vir na tej mrezi pada s certifikatno
    napako (0x8A15005E - "The server certificate did not match any of
    the expected values"), najverjetneje zaradi SSL/TLS inspekcije na
    firewallu/proxyju/AV. Vseh 5 aplikacij spodaj je na voljo v winget
    viru, zato msstore sploh ni potreben.
#>

param(
    [switch]$ResumedByTask,
    [switch]$IncludeDrivers,
    [switch]$SkipApps,
    [switch]$SkipScan,
    [switch]$Reset,
    [switch]$KeepLog,
    [string]$LauncherPath   # pot do .bat datoteke (poda jo batch del)
)

$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'
$ConfirmPreference     = 'None'

# =========================================================================
# KONFIGURACIJA / VAROVALKE
# =========================================================================
$Cfg = @{
    MaxWuPasses        = 8        # prehodov Windows Update skupaj (cez vse restarte)
    MaxDriverPasses    = 3
    MaxReboots         = 8
    MaxRuns            = 14       # zagonov skripte skupaj (zascita proti zanki prijav)
    MaxTotalHours      = 8
    MaxNoProgress      = 3        # zaporednih prehodov brez napredka
    SameSigLimit       = 2        # isti nabor Nx brez namestitve -> izloci ga
    MinFreeGBForWU     = 12
    WuPassTimeoutSec   = 1800     # 30 min na prehod (optimizirano za hitro odkritje zataknjenih WU)
    BootstrapTimeoutSec= 900
    RepairTimeoutSec   = 900
}

$appList = @(
    # Msi = rezervni uradni MSI proizvajalca, ce winget pade (npr. hash mismatch 0x8A150011,
    # ko proizvajalec izda novo verzijo, winget katalog pa se ni posodobljen)
    @{ Name = 'Google Chrome';        Id = 'Google.Chrome';               Msi = 'https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi' },
    @{ Name = 'Mozilla Firefox';      Id = 'Mozilla.Firefox';             Msi = 'https://download.mozilla.org/?product=firefox-msi-latest-ssl&os=win64&lang=en-US' },
    @{ Name = 'Adobe Acrobat Reader'; Id = 'Adobe.Acrobat.Reader.64-bit' },
    @{ Name = '7-Zip';                Id = '7zip.7zip' },
    @{ Name = 'Zoom';                 Id = 'Zoom.Zoom';                   Msi = 'https://zoom.us/client/latest/ZoomInstallerFull.msi?archType=x64' }
)

$taskName    = 'SetupWorkstation-Resume'
$stateDir    = Join-Path $env:ProgramData 'Setup-Workstation'
$stateFile   = Join-Path $stateDir 'state.json'
$logFile     = Join-Path $stateDir 'setup.log'
$workingCopy = Join-Path $stateDir 'Setup-NewWorkstation.ps1'
$lockFile    = Join-Path $stateDir 'running.lock'
$stopFile    = Join-Path $stateDir 'STOP'
$archiveLog  = Join-Path $env:ProgramData 'Setup-Workstation-last.log'

try {
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch { }

# --- GUI Obvestilo funkcija ---
function Show-ScriptAlert {
    param(
        [string]$Message,
        [string]$Title = "Setup Workstation Status",
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        [System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, $Icon) | Out-Null
    } catch { }
}

# --- Zahtevaj Windows PowerShell 5.1 (PS7 nima Checkpoint-Computer) ---
if ($PSVersionTable.PSEdition -eq 'Core') {
    Write-Host "Zaganjam v Windows PowerShell 5.1 (PowerShell 7 nima cmdletov za obnovitveno tocko)..." -ForegroundColor Yellow
    $ps51 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $a = @('-NoExit','-NoProfile','-ExecutionPolicy','Bypass','-File', "`"$PSCommandPath`"")
    if ($IncludeDrivers) { $a += '-IncludeDrivers' }
    if ($SkipApps)       { $a += '-SkipApps' }
    if ($SkipScan)       { $a += '-SkipScan' }
    Start-Process $ps51 -ArgumentList $a
    exit
}

# --- Admin pravice ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Skripta potrebuje admin pravice - ponovni zagon..." -ForegroundColor Yellow
    $fwd = @('-NoExit','-NoProfile','-ExecutionPolicy','Bypass','-File', "`"$PSCommandPath`"")
    if ($IncludeDrivers) { $fwd += '-IncludeDrivers' }
    if ($SkipApps)       { $fwd += '-SkipApps' }
    if ($SkipScan)       { $fwd += '-SkipScan' }
    if ($Reset)          { $fwd += '-Reset' }
    if ($KeepLog)        { $fwd += '-KeepLog' }
    try { Start-Process powershell.exe -Verb RunAs -ArgumentList $fwd }
    catch { Write-Host "Povisanje pravic je bilo zavrnjeno." -ForegroundColor Red; Read-Host "Enter za konec" }
    exit
}

New-Item -ItemType Directory -Path $stateDir -Force | Out-Null

# =========================================================================
# LOG / OSNOVNE FUNKCIJE
# =========================================================================
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    try { Add-Content -LiteralPath $logFile -Value ("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message") -Encoding UTF8 } catch { }
}
function Say {
    param([string]$Message, [string]$Color = 'Gray', [string]$Level = 'INFO')
    Write-Host $Message -ForegroundColor $Color
    Write-Log $Message $Level
}
function Stop-ProcessTree { param([int]$ProcessId) try { & taskkill.exe /PID $ProcessId /T /F 2>&1 | Out-Null } catch { } }
function Remove-Lock { Remove-Item -LiteralPath $lockFile -Force -ErrorAction SilentlyContinue }
function Remove-ResumeTask {
    try { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    try { & schtasks.exe /Delete /TN $taskName /F 2>&1 | Out-Null } catch { }
    try { Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' -Name 'SetupWorkstationResume' -ErrorAction SilentlyContinue } catch { }
}

if ($Reset) {
    Write-Host "Zahtevan -Reset: brisem stanje." -ForegroundColor Yellow
    Remove-Item -LiteralPath $stateFile -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $stopFile  -Force -ErrorAction SilentlyContinue
    Remove-ResumeTask
}

# --- Lock proti vzporednemu zagonu (rocni + task) ---
if (Test-Path -LiteralPath $lockFile) {
    $old = (Get-Content -LiteralPath $lockFile -Raw -ErrorAction SilentlyContinue) -as [int]
    if ($old -and $old -ne $PID -and (Get-Process -Id $old -ErrorAction SilentlyContinue)) {
        Write-Host "Skripta ze tece (PID $old) - koncujem." -ForegroundColor Yellow
        Write-Log "Zavrnjen vzporeden zagon (aktiven PID $old)." 'WARN'
        exit
    }
}
Set-Content -LiteralPath $lockFile -Value $PID -Force -Encoding ASCII

# --- Prvi zagon: prekopiraj se na fiksno lokacijo ---
if (-not (Test-Path -LiteralPath $stateFile) -and $PSCommandPath -and ($PSCommandPath -ine $workingCopy)) {
    Write-Host "Prvi zagon - prestavljam skripto na fiksno lokacijo ($stateDir)..." -ForegroundColor Cyan
    try {
        Copy-Item -LiteralPath $PSCommandPath -Destination $workingCopy -Force
        $srcDir     = Split-Path -Parent $PSCommandPath
        $scriptName = Split-Path -Leaf $PSCommandPath
        $siblingBat = if ($LauncherPath) { $LauncherPath } else {
            Get-ChildItem -Path $srcDir -Filter '*.bat' -ErrorAction SilentlyContinue |
            Where-Object { (Get-Content -LiteralPath $_.FullName -Raw -ErrorAction SilentlyContinue) -match [regex]::Escape($scriptName) } |
            Select-Object -First 1 -ExpandProperty FullName }

        ([PSCustomObject]@{
            OriginalScriptPath = $PSCommandPath
            OriginalBatPath    = $siblingBat
            RunAsUser          = "$env:USERDOMAIN\$env:USERNAME"
            IncludeDrivers     = [bool]$IncludeDrivers
        } | ConvertTo-Json) | Set-Content -LiteralPath (Join-Path $stateDir 'bootstrap.json') -Force -Encoding UTF8

        Write-Log "Skripta prestavljena iz '$PSCommandPath' na '$workingCopy'."
        $a2 = @('-NoExit','-NoProfile','-ExecutionPolicy','Bypass','-File', "`"$workingCopy`"")
        if ($IncludeDrivers) { $a2 += '-IncludeDrivers' }
        if ($SkipApps)       { $a2 += '-SkipApps' }
        if ($SkipScan)       { $a2 += '-SkipScan' }
        if ($KeepLog)        { $a2 += '-KeepLog' }
        Remove-Lock
        Start-Process powershell.exe -ArgumentList $a2
        exit
    } catch {
        Write-Host "Prestavitev ni uspela ($_) - nadaljujem iz trenutne lokacije." -ForegroundColor Yellow
        Write-Log "Prestavitev na fiksno lokacijo NEUSPESNA: $_" 'WARN'
    }
}
if ($PSCommandPath -and ($PSCommandPath -ine $workingCopy) -and -not (Test-Path -LiteralPath $workingCopy)) {
    try { Copy-Item -LiteralPath $PSCommandPath -Destination $workingCopy -Force; Write-Log "Delovna kopija obnovljena." } catch { }
}

# =========================================================================
# STANJE
# =========================================================================
function New-DefaultState {
    [PSCustomObject]@{
        SchemaVersion      = 4
        Runs               = 0
        StartedAt          = (Get-Date -Format 'o')
        PrepDone           = $false
        RestorePointDone   = $false
        WindowsUpdateDone  = $false
        DriverUpdateDone   = $false
        IncludeDrivers     = $false
        AppsDone           = $false
        PdfDefaultDone     = $false
        SystemScanDone     = $false
        WuPass             = 0
        DriverPass         = 0
        WuTotalInstalled   = 0
        WuNoProgress       = 0
        WuLastSig          = ''
        WuSameSigCount     = 0
        WuRepaired         = $false
        SkipIDs            = @()
        SkipTitles         = @()
        RebootCount        = 0
        Warnings           = @()
        OriginalScriptPath = $null
        OriginalBatPath    = $null
        RunAsUser          = "$env:USERDOMAIN\$env:USERNAME"
    }
}

$state = $null
if (Test-Path -LiteralPath $stateFile) {
    try { $state = Get-Content -LiteralPath $stateFile -Raw | ConvertFrom-Json } catch { Write-Log "state.json neberljiv - zacenjam na novo." 'WARN' }
}
if (-not $state) {
    $state = New-DefaultState
    $bsFile = Join-Path $stateDir 'bootstrap.json'
    if (Test-Path -LiteralPath $bsFile) {
        try {
            $bs = Get-Content -LiteralPath $bsFile -Raw | ConvertFrom-Json
            $state.OriginalScriptPath = $bs.OriginalScriptPath
            $state.OriginalBatPath    = $bs.OriginalBatPath
            $state.RunAsUser          = $bs.RunAsUser
            $state.IncludeDrivers     = [bool]$bs.IncludeDrivers
        } catch { }
    }
}
$def = New-DefaultState
foreach ($p in $def.PSObject.Properties) {
    if (@($state.PSObject.Properties.Name) -notcontains $p.Name) {
        $state | Add-Member -NotePropertyName $p.Name -NotePropertyValue $p.Value -Force
    }
}
if ($IncludeDrivers) { $state.IncludeDrivers = $true }
$state.Runs = [int]$state.Runs + 1

function Save-State {
    try { $state | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $stateFile -Force -Encoding UTF8 }
    catch { Write-Log "Save-State ni uspel: $_" 'ERROR' }
}
function Note {
    param([string]$Text)
    $state.Warnings = @($state.Warnings) + $Text
    Write-Log "OPOZORILO: $Text" 'WARN'
    Save-State
}
Save-State

# =========================================================================
# VAROVALKE PROTI ZANKAM
# =========================================================================
function Test-CycleGuard {
    param([string]$Stage = 'SOFTWARE')
    if (Test-Path -LiteralPath $stopFile)                       { return 'STOP_FILE' }
    if ([int]$state.Runs        -gt $Cfg.MaxRuns)               { return 'MAX_RUNS' }
    if ([int]$state.RebootCount -ge $Cfg.MaxReboots)            { return 'MAX_REBOOTS' }
    if ($Stage -eq 'SOFTWARE' -and [int]$state.WuPass     -ge $Cfg.MaxWuPasses)     { return 'MAX_PASSES' }
    if ($Stage -eq 'DRIVERS'  -and [int]$state.DriverPass -ge $Cfg.MaxDriverPasses) { return 'MAX_PASSES' }
    if ([int]$state.WuNoProgress -ge $Cfg.MaxNoProgress)        { return 'NO_PROGRESS' }
    try {
        if (((Get-Date) - [datetime]::Parse($state.StartedAt)).TotalHours -gt $Cfg.MaxTotalHours) { return 'TIME_BUDGET' }
    } catch { }
    return 'OK'
}
function Get-CycleReasonText {
    param([string]$Reason)
    switch ($Reason) {
        'STOP_FILE'   { "rocna zaustavitev (datoteka STOP v $stateDir)" }
        'MAX_RUNS'    { "preseceno stevilo zagonov ($($Cfg.MaxRuns)) - sum na zanko ob prijavi" }
        'MAX_REBOOTS' { "preseceno stevilo avtomatskih restartov ($($Cfg.MaxReboots))" }
        'MAX_PASSES'  { "preseceno stevilo prehodov Windows Update" }
        'NO_PROGRESS' { "$($Cfg.MaxNoProgress) prehodi zapored brez napredka" }
        'TIME_BUDGET' { "presezen casovni proracun ($($Cfg.MaxTotalHours) h)" }
        default       { $Reason }
    }
}

# Ce smo ze cez limit zagonov -> razorozi resume takoj
$skipEverything = $false
if ([int]$state.Runs -gt $Cfg.MaxRuns) {
    Say "!!! Skripta se je zagnala ze $($state.Runs)-krat - to kaze na zanko. Odstranjujem avtomatsko nadaljevanje." 'Red' 'ERROR'
    Remove-ResumeTask
    Note "PREKINJENO: zaznana zanka zagonov ($($state.Runs))."
    Show-ScriptAlert -Message "Skripta se je ustavila! Zaznana je zanka zagonov ($($state.Runs) zagonov)." -Title "Setup Workstation - Napaka Zanke"
    $skipEverything = $true
}

# Resume task odjavi ob vsakem zagonu - ce bo potreben nov restart, se registrira znova
Remove-ResumeTask

$header = if ($ResumedByTask) { 'nadaljevanje po restartu' } elseif ([int]$state.Runs -gt 1) { 'nadaljevanje' } else { 'zacetek' }
Write-Host "`n=== Setup nove delovne postaje - $header ===" -ForegroundColor Cyan
Write-Host "    Zagon #$($state.Runs) | restartov: $($state.RebootCount) | WU prehodov: $($state.WuPass)" -ForegroundColor DarkGray
Write-Host "    Rocna zaustavitev ciklanja: ustvari datoteko $stopFile" -ForegroundColor DarkGray
Write-Log "Zagon #$($state.Runs) (ResumedByTask=$ResumedByTask, WuPass=$($state.WuPass), Reboots=$($state.RebootCount), Drivers=$($state.IncludeDrivers))"

# =========================================================================
# IZOLIRAN ZAGON (locen proces + timeout + zivi izpis)
# =========================================================================
$isoTemplate = @'
$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'
$ConfirmPreference     = 'None'
$IsoArgs = $null
try { $IsoArgs = Get-Content -LiteralPath '__ARGS__' -Raw | ConvertFrom-Json } catch { }
function P { param([string]$t) try { Add-Content -LiteralPath '__PROG__' -Value $t -Encoding UTF8 } catch { } }
$Result = @{ Status = 'UNKNOWN' }
try {
__BODY__
} catch {
    $Result.Status = 'EXCEPTION'
    $Result.Error  = $_.Exception.Message
    P "NAPAKA: $($_.Exception.Message)"
} finally {
    try { $Result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath '__RES__' -Encoding UTF8 } catch { }
}
'@

function Invoke-Isolated {
    param(
        [Parameter(Mandatory)][string]$ScriptText,
        [int]$TimeoutSec = 3600,
        [string]$JobName = 'Job',
        [hashtable]$Arguments = @{}
    )
    $id       = [guid]::NewGuid().ToString('N')
    $tmpPs1   = Join-Path $stateDir "iso-$id.ps1"
    $argsFile = Join-Path $stateDir "iso-$id.args.json"
    $resFile  = Join-Path $stateDir "iso-$id.res.json"
    $progFile = Join-Path $stateDir "iso-$id.prog.txt"
    $outFile  = Join-Path $stateDir "iso-$id.out.txt"
    $errFile  = Join-Path $stateDir "iso-$id.err.txt"
    $allFiles = @($tmpPs1,$argsFile,$resFile,$progFile,$outFile,$errFile)

    $fail = { param($s) [PSCustomObject]@{ TimedOut = $false; Status = $s } }

    $wrapped = $isoTemplate.Replace('__ARGS__', $argsFile).
                            Replace('__PROG__', $progFile).
                            Replace('__RES__',  $resFile).
                            Replace('__BODY__', $ScriptText)
    try {
        ($Arguments | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $argsFile -Encoding UTF8 -ErrorAction Stop
        Set-Content -LiteralPath $tmpPs1 -Value $wrapped -Encoding UTF8 -ErrorAction Stop
        New-Item -ItemType File -Path $progFile -Force | Out-Null
    } catch {
        Say "  '$JobName' - zacasnih datotek ni bilo mogoce pripraviti: $($_.Exception.Message)" 'Red' 'ERROR'
        return (& $fail 'PREP_FAILED')
    }

    $proc = $null
    try {
        $proc = Start-Process powershell.exe -PassThru -WindowStyle Hidden -ArgumentList @(
            '-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File', "`"$tmpPs1`""
        ) -RedirectStandardOutput $outFile -RedirectStandardError $errFile
    } catch {
        Say "  '$JobName' - procesa ni bilo mogoce zagnati: $($_.Exception.Message)" 'Red' 'ERROR'
        Remove-Item -LiteralPath $allFiles -Force -ErrorAction SilentlyContinue
        return (& $fail 'START_FAILED')
    }

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $shown = 0; $lastBeat = -99; $beatShown = $false
    while (-not $proc.HasExited) {
        Start-Sleep -Milliseconds 500
        $lines = @(Get-Content -LiteralPath $progFile -ErrorAction SilentlyContinue)
        if ($lines.Count -gt $shown) {
            if ($beatShown) { Write-Host ''; $beatShown = $false }
            for ($i = $shown; $i -lt $lines.Count; $i++) {
                Write-Host "      $($lines[$i])" -ForegroundColor DarkGray
                Write-Log "  [$JobName] $($lines[$i])"
            }
            $shown = $lines.Count
        }
        if (($sw.Elapsed.TotalSeconds - $lastBeat) -ge 10) {
            $lastBeat = $sw.Elapsed.TotalSeconds
            Write-Host ("`r      [{0:hh\:mm\:ss}] {1} - se izvaja..." -f $sw.Elapsed, $JobName) -NoNewline -ForegroundColor DarkGray
            $beatShown = $true
        }
        if ($sw.Elapsed.TotalSeconds -gt $TimeoutSec) {
            if ($beatShown) { Write-Host '' }
            Say "  '$JobName' je presegel casovno omejitev ($([math]::Round($TimeoutSec/60)) min) - ubijam proces (PID $($proc.Id))." 'Red' 'WARN'
            Stop-ProcessTree $proc.Id
            Start-Sleep -Seconds 2
            Remove-Item -LiteralPath $allFiles -Force -ErrorAction SilentlyContinue
            return [PSCustomObject]@{ TimedOut = $true; Status = 'TIMEOUT' }
        }
    }
    if ($beatShown) { Write-Host '' }
    try { $proc.WaitForExit() } catch { }
    $lines = @(Get-Content -LiteralPath $progFile -ErrorAction SilentlyContinue)
    for ($i = $shown; $i -lt $lines.Count; $i++) {
        Write-Host "      $($lines[$i])" -ForegroundColor DarkGray
        Write-Log "  [$JobName] $($lines[$i])"
    }

    foreach ($f in @($outFile,$errFile)) {
        if (Test-Path -LiteralPath $f) {
            $c = Get-Content -LiteralPath $f -Raw -ErrorAction SilentlyContinue
            if ($c -and $c.Trim()) { try { Add-Content -LiteralPath $logFile -Value "--- $JobName ($(Split-Path $f -Leaf)) ---`r`n$c" -Encoding UTF8 } catch { } }
        }
    }

    $obj = $null
    if (Test-Path -LiteralPath $resFile) {
        try { $obj = Get-Content -LiteralPath $resFile -Raw | ConvertFrom-Json } catch { }
    }
    Remove-Item -LiteralPath $allFiles -Force -ErrorAction SilentlyContinue
    if (-not $obj) { return (& $fail 'NO_RESULT') }
    $obj | Add-Member -NotePropertyName TimedOut -NotePropertyValue $false -Force
    return $obj
}

# =========================================================================
# PENDING REBOOT (trdi vs mehki razlogi)
# =========================================================================
function Get-PendingRebootInfo {
    $hard = @(); $soft = @()
    $cbs = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing'
    if (Test-Path "$cbs\RebootPending")    { $hard += 'CBS RebootPending' }
    if (Test-Path "$cbs\RebootInProgress") { $hard += 'CBS RebootInProgress' }
    if (Test-Path "$cbs\PackagesPending")  { $soft += 'CBS PackagesPending' }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $hard += 'WindowsUpdate RebootRequired' }
    try {
        $a = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName' -Name ComputerName -ErrorAction Stop).ComputerName
        $c = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName'       -Name ComputerName -ErrorAction Stop).ComputerName
        if ($a -and $c -and ($a -ne $c)) { $hard += 'Preimenovanje racunalnika' }
    } catch { }
    try {
        $si = New-Object -ComObject Microsoft.Update.SystemInfo
        if ($si.RebootRequired) { $hard += 'WUA RebootRequired' }
    } catch { }
    try {
        $p = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction Stop).PendingFileRenameOperations
        if ($p) { $soft += 'PendingFileRenameOperations' }
    } catch { }
    [PSCustomObject]@{ Required = ($hard.Count -gt 0); Hard = $hard; Soft = $soft }
}

# =========================================================================
# RESUME TASK + RESTART (z registrom HKLM RunOnce kot rezervno potjo)
# =========================================================================
function Register-ResumeTask {
    $runAsUser = [string]$state.RunAsUser
    $argLine   = "-NoExit -NoProfile -ExecutionPolicy Bypass -File `"$workingCopy`" -ResumedByTask"
    $taskSuccess = $false

    # 1) Primarno: Scheduled Task ob prijavi
    try {
        $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argLine
        $trigger   = New-ScheduledTaskTrigger -AtLogOn -User $runAsUser
        try { $trigger.Delay = 'PT1M' } catch { }
        $principal = New-ScheduledTaskPrincipal -UserId $runAsUser -LogonType Interactive -RunLevel Highest
        $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                        -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit ([TimeSpan]::Zero)
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
            -Principal $principal -Settings $settings -Force -ErrorAction Stop | Out-Null
        if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
            Write-Log "Resume task registriran (AtLogOn, Interactive, $runAsUser)."
            $taskSuccess = $true
        }
    } catch { Write-Log "AtLogOn registracija ni uspela: $($_.Exception.Message)" 'WARN' }

    # 2) Rezerva: Scheduled Task ob zagonu kot SYSTEM
    if (-not $taskSuccess) {
        try {
            $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument ($argLine -replace '-NoExit ', '')
            $trigger   = New-ScheduledTaskTrigger -AtStartup
            try { $trigger.Delay = 'PT2M' } catch { }
            $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
            $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                            -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit ([TimeSpan]::Zero)
            Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
                -Principal $principal -Settings $settings -Force -ErrorAction Stop | Out-Null
            Write-Log 'Resume task registriran (AtStartup, SYSTEM).'
            $taskSuccess = $true
        } catch { Write-Log "SYSTEM registracija ni uspela: $($_.Exception.Message)" 'WARN' }
    }

    # 3) Dodatna rezervna pot: HKLM RunOnce register kljuc
    try {
        $runOnceCmd = "powershell.exe -NoExit -NoProfile -ExecutionPolicy Bypass -File `"$workingCopy`" -ResumedByTask"
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' -Name 'SetupWorkstationResume' -Value $runOnceCmd -Force -ErrorAction SilentlyContinue
        Write-Log 'HKLM RunOnce registrski kljuc uspesno nastavljen kot dodatna rezerva.'
    } catch { Write-Log "RunOnce vpis ni uspel: $($_.Exception.Message)" 'WARN' }

    return ($taskSuccess -or (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' -Name 'SetupWorkstationResume' -ErrorAction SilentlyContinue))
}

function Invoke-SetupRestart {
    param([string]$Reason)
    $g = Test-CycleGuard
    if ($g -ne 'OK') {
        $reasonText = Get-CycleReasonText $g
        Say "`n  Restart bi bil potreben, a ga NE izvedem: $reasonText." 'Red' 'ERROR'
        Note "Restart preskocen ($g)."
        Show-ScriptAlert -Message "Restart preprecen: $reasonText" -Title "Setup Workstation - Restart Preprecen"
        return $false
    }
    Say "`n  Potreben je restart ($Reason). Nastavljam avtomatsko nadaljevanje..." 'Yellow'
    if (-not (Register-ResumeTask)) {
        Say "  Avtomatskega nadaljevanja NI bilo mogoce nastaviti - racunalnika NE restartam." 'Red' 'ERROR'
        Say "  Restartaj rocno, nato znova zazeni: $workingCopy" 'Yellow'
        Note 'Resume task ali RunOnce ni bilo mogoce registrirati - potreben rocen restart.'
        Show-ScriptAlert -Message "Avtomatskega nadaljevanja ni bilo mogoce nastaviti. Prosimo, restartajte rocno." -Title "Setup Workstation - Napaka"
        return $false
    }
    $state.RebootCount = [int]$state.RebootCount + 1
    Save-State
    Say "  Restart #$($state.RebootCount)/$($Cfg.MaxReboots). Za preklic: Ctrl+C ali 'shutdown /a'." 'Yellow'
    for ($i = 20; $i -gt 0; $i--) {
        Write-Host "`r  Restart cez $i s ...   " -NoNewline -ForegroundColor Yellow
        Start-Sleep -Seconds 1
    }
    Write-Host ''
    Remove-Lock
    & shutdown.exe /r /t 5 /c 'Setup delovne postaje - nadaljevanje po restartu'
    Start-Sleep -Seconds 60
    return $true
}

# =========================================================================
# WINDOWS UPDATE - otroski skripti
# =========================================================================
$WuRepairScript = @'
$Result.Status = 'REPAIRED'
$Result.Steps  = @()
P "Popravljam Windows Update sklad..."
foreach ($s in @('wuauserv','bits','cryptsvc','msiserver','trustedinstaller','usosvc','dosvc')) {
    try {
        $svc = Get-Service -Name $s -ErrorAction Stop
        if ($svc.StartType -eq 'Disabled') {
            Set-Service -Name $s -StartupType Manual -ErrorAction SilentlyContinue
            $Result.Steps += "$s : bil onemogocen -> Manual"
            P "  $s : onemogocen -> Manual"
        }
    } catch { }
}
foreach ($s in @('wuauserv','bits','cryptsvc','usosvc')) { try { Stop-Service -Name $s -Force -ErrorAction SilentlyContinue } catch { } }
Start-Sleep -Seconds 3
$stamp = Get-Date -Format 'yyyyMMddHHmmss'
foreach ($p in @("$env:SystemRoot\SoftwareDistribution", "$env:SystemRoot\System32\catroot2")) {
    if (Test-Path -LiteralPath $p) {
        try {
            Rename-Item -LiteralPath $p -NewName ("{0}.setup-{1}" -f (Split-Path $p -Leaf), $stamp) -Force -ErrorAction Stop
            $Result.Steps += "preimenovan: $p"
            P "  preimenovan: $p"
        } catch { $Result.Steps += "NI preimenovan ($p): $($_.Exception.Message)"; P "  NI preimenovan: $p" }
    }
}
foreach ($s in @('cryptsvc','bits','wuauserv')) { try { Start-Service -Name $s -ErrorAction SilentlyContinue } catch { } }
try { & "$env:SystemRoot\System32\UsoClient.exe" StartScan 2>&1 | Out-Null } catch { }
Start-Sleep -Seconds 5
P "Popravilo koncano."
'@

$WuPassScript = @'
$Result.Status         = 'UNKNOWN'
$Result.Found          = 0
$Result.FoundTotal     = 0
$Result.Installed      = 0
$Result.Failed         = 0
$Result.RebootRequired = $false
$Result.Signature      = ''
$Result.PendingIDs     = @()
$Result.Titles         = @()
$Result.FailedTitles   = @()
$Result.PendingTitles  = @()
$Result.Errors         = @()

$mode    = [string]$IsoArgs.Mode
$skipIDs = @($IsoArgs.SkipIDs)

P "Odpiram Windows Update Agent (COM)..."
$session = New-Object -ComObject Microsoft.Update.Session
$session.ClientApplicationID = 'Setup-Workstation'

try {
    $sm = New-Object -ComObject Microsoft.Update.ServiceManager
    $sm.ClientApplicationID = 'Setup-Workstation'
    $sm.AddService2('7971f918-a847-4430-9279-4a52d1efe18d', 7, '') | Out-Null
    P "Microsoft Update (Office, gonilniki) registriran."
} catch { $Result.Errors += "Microsoft Update ni bilo mogoce registrirati: $($_.Exception.Message)" }

$criteria = "IsInstalled=0 and IsHidden=0"
$searcher = $session.CreateUpdateSearcher()
try { $searcher.Online = $true } catch { }

P "Iscem posodobitve ($mode)..."
$sr = $null
try {
    $sr = $searcher.Search($criteria)
} catch {
    $first = $_.Exception.Message
    P "Privzeti vir ni odgovoril - poskusam mimo WSUS na javni Windows Update..."
    try {
        $s2 = $session.CreateUpdateSearcher()
        $s2.ServerSelection = 2
        $sr = $s2.Search($criteria)
        $Result.Errors += "Privzeti vir (WSUS?) ni odgovoril, uporabljen javni Windows Update. Prva napaka: $first"
    } catch {
        $Result.Status = 'SEARCH_FAILED'
        $Result.Error  = "$first | mimo WSUS: $($_.Exception.Message)"
        P "Iskanje NI uspelo."
        return
    }
}

$Result.FoundTotal = [int]$sr.Updates.Count
P ("Skupaj cakajocih posodobitev: " + $Result.FoundTotal)

$ssu  = New-Object -ComObject Microsoft.Update.UpdateColl
$rest = New-Object -ComObject Microsoft.Update.UpdateColl
$ids  = @()

foreach ($u in $sr.Updates) {
    $uid      = [string]$u.Identity.UpdateID
    $isDriver = ($u.Type -eq 2)
    if ($mode -eq 'DRIVERS' -and -not $isDriver) { continue }
    if ($mode -ne 'DRIVERS' -and $isDriver)      { continue }
    if ($skipIDs -contains $uid) { $Result.Errors += "Preskocen (izlocitveni seznam): $($u.Title)"; continue }
    if ($u.InstallationBehavior.CanRequestUserInput) {
        $Result.Errors += "Preskocen (zahteva interakcijo): $($u.Title)"
        continue
    }
    if (-not $u.EulaAccepted) { try { $u.AcceptEula() } catch { } }
    $ids += ("{0}.{1}" -f $uid, $u.Identity.RevisionNumber)
    $Result.PendingIDs    += $uid
    $Result.PendingTitles += [string]$u.Title
    if ($u.Title -match 'Servicing Stack|Update Stack') { $ssu.Add($u) | Out-Null } else { $rest.Add($u) | Out-Null }
}

$Result.Found     = $ssu.Count + $rest.Count
$Result.Signature = (($ids | Sort-Object) -join ';')

if ($Result.Found -eq 0) { $Result.Status = 'NO_UPDATES'; P "Za ta nacin ($mode) ni posodobitev."; return }

P ("Za namestitev: " + $Result.Found)
foreach ($t in $Result.PendingTitles) { P ("  - " + $t) }

function Invoke-Batch {
    param($Coll, [string]$Label)
    $r = @{ Installed = 0; Failed = 0; Reboot = $false; Titles = @(); FailedTitles = @(); Errors = @() }
    if ($null -eq $Coll -or $Coll.Count -eq 0) { return $r }

    P ("$Label - prenasam (" + $Coll.Count + ") ...")
    try {
        $dl = $session.CreateUpdateDownloader()
        $dl.Updates = $Coll
        try { $dl.Priority = 3 } catch { }
        $null = $dl.Download()
    } catch { $r.Errors += "$Label - prenos: $($_.Exception.Message)"; P "$Label - napaka pri prenosu: $($_.Exception.Message)" }

    $ready = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($u in $Coll) {
        if ($u.IsDownloaded) { $ready.Add($u) | Out-Null }
        else { $r.FailedTitles += "$($u.Title) (ni preneseno)"; $r.Failed++ }
    }
    if ($ready.Count -eq 0) { P "$Label - nic ni bilo preneseno."; return $r }

    P ("$Label - namescam (" + $ready.Count + ") ...")
    $ir = $null
    for ($try = 1; $try -le 3; $try++) {
        try {
            $inst = $session.CreateUpdateInstaller()
            $inst.Updates = $ready
            try { $inst.ForceQuiet = $true } catch { }
            if ($inst.RebootRequiredBeforeInstallation) {
                $r.Reboot = $true
                $r.Errors += "$Label - potreben restart PRED namestitvijo."
                P "$Label - potreben restart pred namestitvijo."
                return $r
            }
            $ir = $inst.Install()
            break
        } catch {
            $msg = $_.Exception.Message
            $r.Errors += "$Label - namestitev poskus ${try}: $msg"
            P "$Label - poskus ${try} ni uspel: $msg"
            if ($try -lt 3) { Start-Sleep -Seconds 60 } else { return $r }
        }
    }
    if ($null -eq $ir) { return $r }

    $r.Reboot = [bool]$ir.RebootRequired
    for ($i = 0; $i -lt $ready.Count; $i++) {
        $t = $ready.Item($i).Title
        try {
            $ur = $ir.GetUpdateResult($i)
            if ($ur.ResultCode -eq 2) { $r.Installed++; $r.Titles += $t; P ("  OK   " + $t) }
            else {
                $r.Failed++
                $line = ("{0} [rc={1} hr=0x{2:X8}]" -f $t, $ur.ResultCode, $ur.HResult)
                $r.FailedTitles += $line
                P ("  FAIL " + $line)
            }
        } catch { $r.Failed++; $r.FailedTitles += $t }
    }
    return $r
}

$a = Invoke-Batch -Coll $ssu -Label 'ServicingStack'
if ($a.Installed -gt 0 -and $a.Reboot) {
    $Result.Installed      = $a.Installed
    $Result.Titles         = $a.Titles
    $Result.Errors        += $a.Errors
    $Result.RebootRequired = $true
    $Result.Status         = 'INSTALLED'
    P "Servicing Stack namescen - potreben restart pred ostalimi posodobitvami."
    return
}

$b = Invoke-Batch -Coll $rest -Label 'Posodobitve'

$Result.Installed      = $a.Installed + $b.Installed
$Result.Failed         = $a.Failed + $b.Failed
$Result.Titles         = @($a.Titles) + @($b.Titles)
$Result.FailedTitles   = @($a.FailedTitles) + @($b.FailedTitles)
$Result.Errors        += @($a.Errors) + @($b.Errors)
$Result.RebootRequired = ($a.Reboot -or $b.Reboot)

if ($Result.Installed -gt 0 -and $Result.Failed -eq 0) { $Result.Status = 'INSTALLED' }
elseif ($Result.Installed -gt 0)                       { $Result.Status = 'PARTIAL' }
elseif ($Result.Found -gt 0)                           { $Result.Status = 'INSTALL_FAILED' }
else                                                   { $Result.Status = 'NO_UPDATES' }
P ("Namescenih: " + $Result.Installed + ", neuspelih: " + $Result.Failed)
'@

$WuFallbackScript = @'
$Result.Status = 'UNKNOWN'
P "Rezervna pot: PSWindowsUpdate..."
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    P "Namescam NuGet + PSWindowsUpdate (potreben internet)..."
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false -ErrorAction Stop | Out-Null
    }
    if ((Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue).InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
    }
    Install-Module PSWindowsUpdate -Force -Confirm:$false -Scope AllUsers -SkipPublisherCheck -ErrorAction Stop
}
Import-Module PSWindowsUpdate -ErrorAction Stop
P "Namescam posodobitve preko PSWindowsUpdate..."
$res = Get-WindowsUpdate -Install -AcceptAll -IgnoreReboot -Confirm:$false -ErrorAction Stop
$Result.Installed = @($res).Count
$Result.Status    = if (@($res).Count -gt 0) { 'INSTALLED' } else { 'NO_UPDATES' }
try { $Result.RebootRequired = [bool](Get-WURebootStatus -Silent) } catch { }
P ("Rezervna pot koncana: " + $Result.Status)
'@

# =========================================================================
# CIKLICNI WU PREHODI (z avtomatskim resetom wuauserv ob zataknitvi)
# =========================================================================
function Invoke-UpdateStage {
    param(
        [ValidateSet('SOFTWARE','DRIVERS')][string]$Mode,
        [string]$PassProp,
        [int]$MaxPasses,
        [string]$Label
    )
    while ($true) {
        $g = Test-CycleGuard -Stage $Mode
        if ($g -ne 'OK') {
            $reasonText = Get-CycleReasonText $g
            Say "`n  CIKLANJE USTAVLJENO: $reasonText" 'Red' 'WARN'
            Note "$Label ustavljen ($g) po $($state.$PassProp) prehodih, $($state.WuTotalInstalled) namescenih posodobitvah."
            Show-ScriptAlert -Message "Windows Update ustavljen: $reasonText" -Title "Setup Workstation - WU Ustavljen"
            return $false
        }

        $state.$PassProp = [int]$state.$PassProp + 1
        Save-State
        Say "`n  >> $Label - prehod $($state.$PassProp)/$MaxPasses" 'Cyan'

        $wu = Invoke-Isolated -ScriptText $WuPassScript -TimeoutSec $Cfg.WuPassTimeoutSec `
                -JobName "WU-$Mode-pass$($state.$PassProp)" `
                -Arguments @{ Mode = $Mode; SkipIDs = @($state.SkipIDs) }

        $status    = if ($wu.TimedOut) { 'TIMEOUT' } else { [string]$wu.Status }
        $installed = [int]$wu.Installed
        Say "  Status: $status | najdenih: $($wu.Found) | namescenih: $installed | neuspelih: $($wu.Failed)" 'Gray'
        if ($wu.Errors) { foreach ($e in @($wu.Errors)) { Say "    ! $e" 'Yellow' 'WARN' } }
        $state.WuTotalInstalled = [int]$state.WuTotalInstalled + $installed

        # ---- Odpoved ali zataknitev -> ponovni zagon storitve wuauserv & popravilo sklada ----
        if ($status -in @('SEARCH_FAILED','INSTALL_FAILED','EXCEPTION','TIMEOUT','NO_RESULT','START_FAILED','PREP_FAILED','UNKNOWN')) {
            if ($wu.Error) { Say "    Napaka: $($wu.Error)" 'Red' 'ERROR' }

            Say "  Zaznana zataknitev ali neuspeh ($status). Zaganjam reset storitve Windows Update (wuauserv)..." 'Yellow'
            try {
                Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 5
                Start-Service -Name wuauserv -ErrorAction SilentlyContinue
            } catch { }

            if (-not [bool]$state.WuRepaired) {
                $state.WuRepaired = $true; Save-State
                Say "  Poskusam popraviti Windows Update sklad (servisi + SoftwareDistribution/catroot2)..." 'Yellow'
                $rep = Invoke-Isolated -ScriptText $WuRepairScript -TimeoutSec $Cfg.RepairTimeoutSec -JobName 'WU-Repair'
                if ($rep.Steps) { foreach ($s in @($rep.Steps)) { Write-Log "  * $s" } }
                Note 'Windows Update sklad je bil popravljen (resetiran SoftwareDistribution/catroot2).'
                Start-Sleep -Seconds 10
                continue
            }

            Say "  COM pot je odpovedala tudi po popravilu - poskusam PSWindowsUpdate (rezerva)..." 'Yellow'
            $fb = Invoke-Isolated -ScriptText $WuFallbackScript -TimeoutSec $Cfg.WuPassTimeoutSec -JobName 'WU-Fallback'
            if ($fb.Status -in @('INSTALLED','NO_UPDATES')) {
                Say "  Rezervna pot: $($fb.Status), namescenih $($fb.Installed)." 'Green'
                $state.WuTotalInstalled = [int]$state.WuTotalInstalled + [int]$fb.Installed
                Save-State
                if ([bool]$fb.RebootRequired -or (Get-PendingRebootInfo).Required) {
                    if (Invoke-SetupRestart -Reason "$Label (rezervna pot)") { exit }
                    return $false
                }
                return $true
            }
            Say "  $Label NI uspel. Potreben je rocen pregled." 'Red' 'ERROR'
            Note "$Label NI USPEL ($status). Preveri: Nastavitve > Windows Update in %SystemRoot%\Logs\CBS\CBS.log."
            Show-ScriptAlert -Message "Windows Update ni uspel ($status). Prosimo, preverite sistemski dnevnik." -Title "Setup Workstation - WU Napaka"
            return $false
        }

        # ---- Ni vec posodobitev ----
        if ($status -eq 'NO_UPDATES') {
            $pr = Get-PendingRebootInfo
            if ($pr.Required) {
                Say "  Ni vec posodobitev, a sistem ceka restart ($($pr.Hard -join ', '))." 'Yellow'
                if (Invoke-SetupRestart -Reason "$Label - zakljucni restart") { exit }
                return $false
            }
            Say "  $Label je CIST - ni cakajocih posodobitev." 'Green'
            return $true
        }

        # ---- Zaznavanje ciklanja ----
        $sig = [string]$wu.Signature
        if ($installed -eq 0) {
            if ($sig -and $sig -eq [string]$state.WuLastSig) {
                $state.WuSameSigCount = [int]$state.WuSameSigCount + 1
                Say "  Isti nabor posodobitev kot prejsnji prehod ($($state.WuSameSigCount)/$($Cfg.SameSigLimit))." 'Red' 'WARN'
                if ($state.WuSameSigCount -ge $Cfg.SameSigLimit) {
                    Say "  CIKLANJE ZAZNANO - te posodobitve izlocam iz nadaljnjih prehodov:" 'Red' 'WARN'
                    foreach ($t in @($wu.PendingTitles)) { Say "    x $t" 'Red' 'WARN' }
                    $state.SkipIDs        = @($state.SkipIDs) + @($wu.PendingIDs)
                    $state.SkipTitles     = @($state.SkipTitles) + @($wu.PendingTitles)
                    $state.WuSameSigCount = 0
                    $state.WuNoProgress   = 0
                    $state.WuLastSig      = ''
                    Note "Izlocene posodobitve (ciklanje): $((@($wu.PendingTitles)) -join ' | ')"
                    Save-State
                    continue
                }
            } else {
                $state.WuSameSigCount = 1
            }
            $state.WuNoProgress = [int]$state.WuNoProgress + 1
            Say "  Brez napredka ($($state.WuNoProgress)/$($Cfg.MaxNoProgress))." 'Yellow' 'WARN'
        } else {
            $state.WuNoProgress   = 0
            $state.WuSameSigCount = 0
        }
        $state.WuLastSig = $sig
        Save-State

        # ---- Restart, ce je potreben ----
        $pr = Get-PendingRebootInfo
        if ([bool]$wu.RebootRequired -or $pr.Required) {
            if (Invoke-SetupRestart -Reason "$Label prehod $($state.$PassProp)") { exit }
            return $false
        }
        Say "  Restart ni potreben - grem v naslednji prehod." 'DarkGray'
        Start-Sleep -Seconds 5
    }
}

# =========================================================================
# 1. PREDPRIPRAVA
# =========================================================================
$fatal = $null
try {
if (-not $skipEverything) {

Write-Host "`n[1/7] Predpriprava (poraba energije, Fast Startup, servisi, preverjanja)" -ForegroundColor Cyan
if (-not $state.PrepDone) {
    try { & powercfg.exe /setactive SCHEME_MIN 2>&1 | Out-Null } catch { }
    foreach ($t in @('monitor-timeout-ac 0','monitor-timeout-dc 0','standby-timeout-ac 0','standby-timeout-dc 0','hibernate-timeout-ac 0','hibernate-timeout-dc 0')) {
        try { Start-Process powercfg.exe -ArgumentList "/change $t" -NoNewWindow -Wait } catch { }
    }
    Say "  Power plan: brez spanja in ugasanja zaslona (AC + baterija)." 'Gray'

    # --- Onemogoci Fast Startup za zanesljivo izvajanje ob restartu ---
    try {
        Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' -Name 'HiberbootEnabled' -Value 0 -Force -ErrorAction SilentlyContinue
        Say "  Fast Startup (Hitri zagon): onemogocen za zanesljive restarte." 'Gray'
    } catch { }

    foreach ($svc in @('wuauserv','bits','cryptsvc','msiserver','usosvc','TrustedInstaller')) {
        try {
            $s = Get-Service -Name $svc -ErrorAction Stop
            if ($s.StartType -eq 'Disabled') {
                Set-Service -Name $svc -StartupType Manual -ErrorAction SilentlyContinue
                Say "  Servis $svc je bil onemogocen - nastavljen na Manual." 'Yellow' 'WARN'
            }
            if ($svc -in @('wuauserv','bits','cryptsvc') -and $s.Status -ne 'Running') { Start-Service -Name $svc -ErrorAction SilentlyContinue }
        } catch { Write-Log "Servis $svc ni dostopen: $_" 'WARN' }
    }

    try {
        $wsus = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -ErrorAction SilentlyContinue
        if ($wsus -and $wsus.WUServer) {
            Say "  Politika WSUS: $($wsus.WUServer)" 'Yellow' 'WARN'
            Note "Racunalnik je vezan na WSUS ($($wsus.WUServer)) - ce ni dosegljiv, skripta samodejno uporabi javni Windows Update."
        }
    } catch { }

    try {
        $freeGB = [math]::Round((Get-PSDrive -Name ($env:SystemDrive.TrimEnd(':')) -ErrorAction Stop).Free / 1GB, 1)
        Say "  Prosto na $($env:SystemDrive) : $freeGB GB" 'Gray'
        if ($freeGB -lt $Cfg.MinFreeGBForWU) {
            Say "  OPOZORILO: premalo prostora (priporoceno vsaj $($Cfg.MinFreeGBForWU) GB) - WU lahko pade z 0x80070070." 'Red' 'WARN'
            Note "Premalo prostora ($freeGB GB) za zanesljiv Windows Update."
        }
    } catch { }

    $state.PrepDone = $true
    Save-State
} else { Write-Host "      Ze opravljeno - preskocim." -ForegroundColor DarkGray }

# =========================================================================
# 2. OBNOVITVENA TOCKA
# =========================================================================
Write-Host "`n[2/7] Obnovitvena tocka" -ForegroundColor Cyan
if (-not $state.RestorePointDone) {
    $restoreName = 'naveza-' + (Get-Date -Format 'yyyy-MM-dd')
    $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    if ($osInfo -and $osInfo.Caption -match 'Server') {
        Say "  Windows Server ne podpira System Restore - preskocim." 'DarkGray'
        $state.RestorePointDone = $true
    } else {
        try {
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore' -Name 'SystemRestorePointCreationFrequency' -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
            Enable-ComputerRestore -Drive "$env:SystemDrive\" -ErrorAction Stop
            Checkpoint-Computer -Description $restoreName -RestorePointType 'MODIFY_SETTINGS' -ErrorAction Stop
            Say "  Obnovitvena tocka '$restoreName' ustvarjena." 'Green'
            $state.RestorePointDone = $true
        } catch {
            Say "  Napaka pri obnovitveni tocki: $_" 'Red' 'ERROR'
            Say "  Poskusim znova ob naslednjem zagonu." 'Yellow'
        }
    }
    Save-State
} else { Write-Host "      Ze ustvarjena - preskocim." -ForegroundColor DarkGray }

# =========================================================================
# 3. WINDOWS UPDATE
# =========================================================================
Write-Host "`n[3/7] Windows Update (COM WUA, ciklicno, avtonomno)" -ForegroundColor Cyan
if (-not $state.WindowsUpdateDone) {
    if (Invoke-UpdateStage -Mode 'SOFTWARE' -PassProp 'WuPass' -MaxPasses $Cfg.MaxWuPasses -Label 'Windows Update') {
        $state.WindowsUpdateDone = $true
    }
    Save-State
} else { Write-Host "      Ze opravljen - preskocim." -ForegroundColor DarkGray }

# =========================================================================
# 4. GONILNIKI (OPCIJSKO)
# =========================================================================
Write-Host "`n[4/7] Gonilniki preko Windows Update (opcijsko)" -ForegroundColor Cyan
if (-not $state.IncludeDrivers) {
    Write-Host "      Izklopljeno (vklop: -IncludeDrivers)." -ForegroundColor DarkGray
    $state.DriverUpdateDone = $true; Save-State
} elseif (-not $state.WindowsUpdateDone) {
    Write-Host "      Cakam, da se najprej zakljuci Windows Update." -ForegroundColor DarkGray
} elseif ($state.DriverUpdateDone) {
    Write-Host "      Ze opravljeno - preskocim." -ForegroundColor DarkGray
} else {
    $state.WuNoProgress = 0; Save-State
    if (-not (Invoke-UpdateStage -Mode 'DRIVERS' -PassProp 'DriverPass' -MaxPasses $Cfg.MaxDriverPasses -Label 'Gonilniki')) {
        Note 'Gonilniki niso bili v celoti namesceni (best effort).'
    }
    $state.DriverUpdateDone = $true
    Save-State
}

# =========================================================================
# 5. APLIKACIJE
# =========================================================================
Write-Host "`n[5/7] Aplikacije (winget)" -ForegroundColor Cyan
function Get-WingetPath {
    $cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    try { Add-AppxPackage -RegisterByFamilyName -MainPackage 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe' -ErrorAction SilentlyContinue } catch { }
    $cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $cand = Get-ChildItem "$env:ProgramFiles\WindowsApps" -Filter 'Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe' -Directory -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending | Select-Object -First 1
    if ($cand) { $exe = Join-Path $cand.FullName 'winget.exe'; if (Test-Path $exe) { return $exe } }
    return $null
}
if ($SkipApps) {
    Write-Host "      Preskoceno (-SkipApps)." -ForegroundColor DarkGray
    $state.AppsDone = $true; Save-State
} elseif (-not $state.AppsDone) {
    $winget = Get-WingetPath
    if (-not $winget) {
        Say "  winget ni na voljo v tem kontekstu - poskusim ob naslednjem zagonu (prijavi se kot uporabnik)." 'Yellow' 'WARN'
    } else {
        $wgVer = ''
        try { $wgVer = (& $winget --version) -join '' } catch { }
        Write-Log "winget: $winget ($wgVer)"
        $extra = @()
        if ($wgVer -match 'v?(\d+)\.(\d+)') {
            if ([int]$Matches[1] -gt 1 -or ([int]$Matches[1] -eq 1 -and [int]$Matches[2] -ge 4)) { $extra += '--disable-interactivity' }
        }
        try { & $winget source update 2>&1 | Out-Null } catch { }

        $anyFailed = $false
        foreach ($app in $appList) {
            Write-Host "   -> $($app.Name)..." -ForegroundColor Yellow
            # --source winget: obvoz za msstore cert napako (0x8A15005E - "server certificate did not match")
            $wgArgs = @('install','--id',$app.Id,'-e','--silent','--accept-package-agreements','--accept-source-agreements','--source','winget') + $extra
            try { & $winget @wgArgs; $exitCode = $LASTEXITCODE } catch { Write-Log "winget $($app.Id): $_" 'ERROR'; $exitCode = -1 }
            $ok = $false
            try { & $winget list --id $app.Id -e --accept-source-agreements | Out-Null; $ok = ($LASTEXITCODE -eq 0) } catch { }
            if (-not $ok -and $app.Msi) {
                Say ("      {0}: winget ni uspel (exit {1} / 0x{1:X8}) - poskusam uradni MSI..." -f $app.Name, $exitCode) 'Yellow' 'WARN'
                $msi = Join-Path $env:TEMP ("setup-" + ($app.Id -replace '[^\w\.]','_') + '.msi')
                try {
                    Invoke-WebRequest -Uri $app.Msi -OutFile $msi -UseBasicParsing -ErrorAction Stop
                    $mp = Start-Process msiexec.exe -ArgumentList "/i `"$msi`" /qn /norestart" -Wait -PassThru
                    $exitCode = $mp.ExitCode
                    Write-Log "MSI $($app.Id): exit $exitCode"
                } catch { Write-Log "MSI $($app.Id): $_" 'ERROR'; $exitCode = -1 }
                Remove-Item -LiteralPath $msi -Force -ErrorAction SilentlyContinue
                try { & $winget list --id $app.Id -e --accept-source-agreements | Out-Null; $ok = ($LASTEXITCODE -eq 0) } catch { }
                if (-not $ok -and $exitCode -in @(0, 3010)) { $ok = $true }
            }
            if ($ok) { Say "      $($app.Name): namescen (exit $exitCode)." 'Green' }
            else     { Say ("      {0}: NI namescen (exit {1} / 0x{1:X8})." -f $app.Name, $exitCode) 'Red' 'WARN'; $anyFailed = $true }
        }
        if ($anyFailed) { Say "  En ali vec programov se ni namestilo - poskusim ob naslednjem zagonu." 'Yellow' 'WARN' }
        else            { $state.AppsDone = $true }
    }
    Save-State
} else { Write-Host "      Ze namescene - preskocim." -ForegroundColor DarkGray }

# =========================================================================
# 6. PRIVZETI PDF
# =========================================================================
Write-Host "`n[6/7] Privzeti PDF pregledovalnik" -ForegroundColor Cyan
if (-not $state.PdfDefaultDone) {
    $interactive = $false
    try { $interactive = [Environment]::UserInteractive -and ((Get-Process -Id $PID).SessionId -ne 0) } catch { }
    if ($interactive) {
        try {
            Start-Process 'ms-settings:defaultapps' -ErrorAction SilentlyContinue
            Say "  Odprl sem Nastavitve - PDF privzeti program potrdi rocno (Windows tega ne dovoli tiho)." 'Yellow'
            $state.PdfDefaultDone = $true
        } catch { }
    } else {
        Say "  Ni interaktivne seje - privzeti PDF ostane za rocno potrditev." 'DarkGray'
    }
    Save-State
} else { Write-Host "      Ze urejeno - preskocim." -ForegroundColor DarkGray }

# =========================================================================
# 7. DISM + SFC
# =========================================================================
Write-Host "`n[7/7] DISM RestoreHealth + SFC /scannow" -ForegroundColor Cyan
if ($SkipScan) {
    Write-Host "      Preskoceno (-SkipScan)." -ForegroundColor DarkGray
    $state.SystemScanDone = $true; Save-State
} elseif (-not $state.SystemScanDone) {
    Say "  Zaganjam DISM (lahko traja 10-20 min)..." 'Yellow'
    $t0 = Get-Date
    DISM /Online /Cleanup-Image /RestoreHealth
    $dismRc = $LASTEXITCODE
    Write-Log "DISM exit code $dismRc ($([int]((Get-Date)-$t0).TotalMinutes) min)"
    if ($dismRc -ne 0) { Note "DISM /RestoreHealth ni uspel (exit $dismRc) - component store je lahko poskodovan." }
    Say "  Zaganjam SFC /scannow..." 'Yellow'
    $t0 = Get-Date
    sfc /scannow
    Write-Log "SFC exit code $LASTEXITCODE ($([int]((Get-Date)-$t0).TotalMinutes) min)"
    $state.SystemScanDone = $true
    Save-State
} else { Write-Host "      Ze opravljeno - preskocim." -ForegroundColor DarkGray }

}

} catch {
    $fatal = $_
    Say "`n!!! NEPRICAKOVANA NAPAKA: $($_.Exception.Message)" 'Red' 'ERROR'
    Write-Log $_.ScriptStackTrace 'ERROR'
    Show-ScriptAlert -Message "Nepricakovana napaka pri izvojenju:`n$($_.Exception.Message)" -Title "Setup Workstation - Usodna Napaka"
} finally {
    Remove-ResumeTask
}

# =========================================================================
# POVZETEK
# =========================================================================
$fullySucceeded = $state.WindowsUpdateDone -and $state.RestorePointDone -and $state.AppsDone -and (-not $fatal)

Write-Host "`n----------------------------------------------------" -ForegroundColor Cyan
Write-Host ("  Windows Update : {0}   (prehodov: {1}, namescenih: {2}, restartov: {3})" -f $state.WindowsUpdateDone, $state.WuPass, $state.WuTotalInstalled, $state.RebootCount)
Write-Host ("  Obnov. tocka   : {0}" -f $state.RestorePointDone)
Write-Host ("  Gonilniki      : {0}" -f $(if ($state.IncludeDrivers) { $state.DriverUpdateDone } else { 'izklopljeno' }))
Write-Host ("  Aplikacije     : {0}" -f $state.AppsDone)
Write-Host ("  DISM + SFC     : {0}" -f $state.SystemScanDone)
if (@($state.SkipTitles).Count -gt 0) {
    Write-Host "`n  Izlocene posodobitve (ciklanje / neuspeh):" -ForegroundColor Yellow
    foreach ($t in @($state.SkipTitles)) { Write-Host "    x $t" -ForegroundColor Yellow }
}
if (@($state.Warnings).Count -gt 0) {
    Write-Host "`n  Opozorila:" -ForegroundColor Yellow
    foreach ($w in @($state.Warnings)) { Write-Host "    - $w" -ForegroundColor Yellow }
}
$prEnd = Get-PendingRebootInfo
if ($prEnd.Required) { Write-Host "  !!! SE VEDNO POTREBEN RESTART: $($prEnd.Hard -join ', ')" -ForegroundColor Red }
elseif ($prEnd.Soft.Count) { Write-Host "  (Mehki signal za restart: $($prEnd.Soft -join ', ') - ni nujno.)" -ForegroundColor DarkGray }
Write-Host "----------------------------------------------------" -ForegroundColor Cyan

if ($fullySucceeded) {
    Write-Host "`n=== Setup koncan ===" -ForegroundColor Green
    Write-Host "Rocno preveri se: timezone, ime racunalnika, domain join, tiskalniki, privzeti PDF." -ForegroundColor Cyan
    Write-Log 'Setup uspesno koncan - cistim za seboj.'

    Remove-ResumeTask
    if ($KeepLog -or @($state.Warnings).Count -gt 0) {
        try { Copy-Item -LiteralPath $logFile -Destination $archiveLog -Force -ErrorAction SilentlyContinue
              Write-Host "Dnevnik shranjen: $archiveLog" -ForegroundColor DarkGray } catch { }
    }
    foreach ($p in @($state.OriginalScriptPath, $state.OriginalBatPath)) {
        if ($p -and (Test-Path -LiteralPath $p)) { Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue }
    }
    Remove-Lock
    try {
        Start-Process -WindowStyle Hidden cmd.exe -ArgumentList "/c timeout /t 5 >nul & rmdir /s /q `"$stateDir`""
        Write-Host "Skripta se bo pocistila." -ForegroundColor Green
    } catch { Remove-Item -LiteralPath $stateDir -Recurse -Force -ErrorAction SilentlyContinue }
} else {
    Write-Host "`n=== Setup KONCAN Z OPOZORILI ===" -ForegroundColor Yellow
    Write-Host "Skripta se NAMENOMA ni odstranila - nekaj ni v celoti uspelo." -ForegroundColor Yellow
    Write-Host "Dnevnik: $logFile" -ForegroundColor Yellow
    Write-Host "Znova pozeni: powershell -ExecutionPolicy Bypass -File `"$workingCopy`"" -ForegroundColor Yellow
    Write-Log "Setup koncan Z OPOZORILI. WU=$($state.WindowsUpdateDone) RP=$($state.RestorePointDone) Apps=$($state.AppsDone)" 'WARN'
    Remove-Lock

    $alertMsg = "Setup se je zakljucil Z OPOZORILI ali NAPAKAMI.`n`nDnevnik: $logFile`n`nStatus:`n- Windows Update: $($state.WindowsUpdateDone)`n- Obnovitvena tocka: $($state.RestorePointDone)`n- Aplikacije: $($state.AppsDone)"
    Show-ScriptAlert -Message $alertMsg -Title "Setup Workstation - Z Opozorili"
}
