<#
.SYNOPSIS
  Weekly earnings-calendar refresh: pulls fresh FactSet values through Excel,
  validates them, writes data.json and pushes it to GitHub (GitHub Pages then
  republishes the site).

.DESCRIPTION
  1. Copies the master workbook to a temp folder (the original is never opened
     for writing and never saved).
  2. Opens the copy in a hidden Excel, loads the FactSet add-in and runs
     FDS_RECALC_NOW (the same refresh as the FactSet ribbon / Alt+3).
  3. Reads the "Web Pulls" sheet from row 3 down to the first blank ticker, so
     tickers added to the sheet are picked up automatically.
  4. Validates the result. If validation fails nothing is published and the
     site keeps last week's data.
  5. Writes data.json, commits and pushes.

  Runs fine while the workstation is locked, as long as the user is logged on
  and FactSet is signed in.

.PARAMETER NoPush
  Do everything except commit/push (writes data.json locally).

.PARAMETER Force
  Publish even if the "big change vs last week" check fails.
#>
param(
  [string]$WorkbookPath = '\\archnasuni\G&M-Research\G&M\Admin\Earnings Calendar\Key Earnings Calendar.xlsx',
  [string]$SheetName    = 'Web Pulls',
  [string]$RepoPath     = (Split-Path -Parent $PSScriptRoot),
  [string]$XllPath      = 'C:\Program Files (x86)\FactSet\FDSXL32_x64.XLL',
  [int]$TimeoutSeconds  = 300,
  [switch]$NoPush,
  [switch]$Force,
  # Team email after a successful publish: Send (scheduled runs), Draft (save to
  # Drafts for review), or None.
  [ValidateSet('Send', 'Draft', 'None')][string]$EmailMode = 'Send',
  [string]$EmailTo  = 'GMAnalysts; GMExecAdmin',
  [string]$SiteUrl  = 'https://tylerkw74-dev.github.io/earnings-calendar/'
)

$ErrorActionPreference = 'Stop'
$logDir = Join-Path $env:LOCALAPPDATA 'EarningsCalendar\logs'
New-Item -ItemType Directory -Force $logDir | Out-Null
$logFile = Join-Path $logDir ('refresh-{0:yyyy-MM-dd_HHmmss}.log' -f (Get-Date))
function Log([string]$msg) { $line = '{0:yyyy-MM-dd HH:mm:ss}  {1}' -f (Get-Date), $msg; Add-Content -LiteralPath $logFile -Value $line; Write-Output $line }

function Notify([string]$title, [string]$body) {
  # Windows toast notification (shows in Action Center even if the PC was locked).
  try {
    [void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
    [void][Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
    $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
    $esc = { param($s) [Security.SecurityElement]::Escape($s) }
    $xml.LoadXml("<toast><visual><binding template='ToastGeneric'><text>$(& $esc $title)</text><text>$(& $esc $body)</text></binding></visual></toast>")
    $appId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show([Windows.UI.Notifications.ToastNotification]::new($xml))
  } catch { Log "toast failed: $($_.Exception.Message)" }
}

function Wait-SiteLive([string]$generatedAt, [int]$timeoutSec = 600) {
  # GitHub Pages takes ~1 minute to republish; confirm the live data.json is the new one.
  $deadline = (Get-Date).AddSeconds($timeoutSec)
  do {
    try {
      $live = Invoke-RestMethod -UseBasicParsing -Uri ($SiteUrl + 'data.json?nocache=' + [guid]::NewGuid().ToString('N')) -Headers @{ 'Cache-Control' = 'no-cache' }
      if ($live.generatedAt -eq $generatedAt) { return $true }
    } catch {}
    Start-Sleep -Seconds 20
  } while ((Get-Date) -lt $deadline)
  return $false
}

function Send-TeamEmail([string]$mode) {
  $subject = 'Key Earnings Calendar {0:M.d.yyyy}' -f (Get-Date)
  $html = "<div style=""font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt"">" +
          "<p>Hi Team,</p>" +
          "<p>Please see the updated earnings calendar at this link: <a href=""$SiteUrl"">$SiteUrl</a>.</p>" +
          "<p>Best,<br>Tyler</p></div>"
  $ol = New-Object -ComObject Outlook.Application
  $mail = $ol.CreateItem(0)
  $mail.To = $EmailTo
  $mail.Subject = $subject
  $mail.HTMLBody = $html
  if (-not $mail.Recipients.ResolveAll()) {
    $bad = @($mail.Recipients | Where-Object { -not $_.Resolved } | ForEach-Object Name) -join ', '
    $mail.Save(); throw "Outlook could not resolve recipient(s): $bad (saved to Drafts instead)"
  }
  if ($mode -eq 'Send') { $mail.Send(); Log "Email sent: '$subject' to $EmailTo" }
  else { $mail.Save(); Log "Email saved to Drafts (not sent): '$subject' to $EmailTo" }
}

Add-Type -Namespace Native -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);' -ErrorAction SilentlyContinue

$tempDir = Join-Path $env:TEMP ('earnings-refresh-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$xl = $null; $xlPid = 0
try {
  Log "Start. Workbook: $WorkbookPath  Repo: $RepoPath"
  if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
  New-Item -ItemType Directory -Force $tempDir | Out-Null
  $copy = Join-Path $tempDir 'workbook.xlsx'
  Copy-Item -LiteralPath $WorkbookPath -Destination $copy

  # --- Refresh in a hidden Excel -------------------------------------------
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.AskToUpdateLinks = $false
  [uint32]$p = 0; [void][Native.Win]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$p); $xlPid = $p
  if (-not $xl.RegisterXLL($XllPath)) { throw "Could not load the FactSet add-in ($XllPath)" }
  [void]$xl.Workbooks.Add(); $xl.Calculation = -4135   # manual: nothing recalcs until we ask
  $wb = $xl.Workbooks.Open($copy, 0, $false)
  $ws = $wb.Worksheets.Item($SheetName)

  Log 'Running FDS_RECALC_NOW'
  [void]$xl.Run('FDS_RECALC_NOW')
  $t0 = Get-Date
  $busyPattern = '(?i)#calc|getting|loading|request|waiting'
  do {
    Start-Sleep -Seconds 3
    $last = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row
    $busy = 0
    if ($last -ge 3) {
      foreach ($col in 'B', 'C', 'E', 'H') {
        foreach ($v in @($ws.Range("${col}3:${col}$last").Value2)) { if ([string]$v -match $busyPattern) { $busy++ } }
      }
    }
  } while ($busy -gt 0 -and ((Get-Date) - $t0).TotalSeconds -lt $TimeoutSeconds)
  Log ("Refresh settled after {0:N0}s (cells still loading: {1})" -f ((Get-Date) - $t0).TotalSeconds, $busy)
  if ($busy -gt 0) { throw "FactSet did not finish within $TimeoutSeconds s ($busy cells still loading). Is FactSet signed in?" }

  # --- Read rows ------------------------------------------------------------
  $rows = New-Object System.Collections.Generic.List[object]
  $r = 3
  while ($true) {
    $ticker = ([string]$ws.Range("A$r").Text).Trim()
    if (-not $ticker) { break }
    $rows.Add([pscustomobject]@{
      Ticker  = $ticker
      Date    = ([string]$ws.Range("B$r").Text).Trim()
      Time    = ([string]$ws.Range("C$r").Text).Trim()
      Timing  = ([string]$ws.Range("D$r").Text).Trim()
      Company = ([string]$ws.Range("E$r").Text).Trim()
      CE      = ([string]$ws.Range("H$r").Text).Trim()
    })
    $r++
  }
  $wb.Close($false)
  Log "Read $($rows.Count) tickers from '$SheetName'"

  # --- Validate -------------------------------------------------------------
  $warnings = New-Object System.Collections.Generic.List[string]
  if ($rows.Count -eq 0) { throw 'No tickers found on the sheet.' }
  $dupes = $rows | Group-Object Ticker | Where-Object Count -gt 1
  foreach ($d in $dupes) { $warnings.Add("Duplicate ticker $($d.Name) ($($d.Count) rows) - kept the first") }
  $seen = @{}; $clean = foreach ($row in $rows) { if (-not $seen.ContainsKey($row.Ticker)) { $seen[$row.Ticker] = 1; $row } }
  $dated = @($clean | Where-Object { $_.Date -match '^\d{2}/\d{2}/\d{4} [CE]$' })
  $undated = @($clean | Where-Object { $_.Date -notmatch '^\d{2}/\d{2}/\d{4} [CE]$' })
  foreach ($u in $undated) { $warnings.Add("No date for $($u.Ticker) ($($u.Date))") }
  $pct = [math]::Round(100 * $dated.Count / $clean.Count)
  Log "Valid dates: $($dated.Count)/$($clean.Count) ($pct%)"
  if ($pct -lt 80) { throw "Only $pct% of tickers have a valid date - looks like a failed pull. Not publishing." }

  # Compare with last published data
  $dataPath = Join-Path $RepoPath 'data.json'
  $prev = $null
  if (Test-Path -LiteralPath $dataPath) { try { $prev = Get-Content -LiteralPath $dataPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { $warnings.Add('Previous data.json unreadable') } }
  if ($prev -and $prev.count -gt 0) {
    $drop = 1 - ($dated.Count / [double]$prev.count)
    if ($drop -gt 0.25 -and -not $Force) { throw ("Dated rows fell from {0} to {1} (more than 25%). Not publishing; rerun with -Force if intended." -f $prev.count, $dated.Count) }
    $prevTickers = @(($prev.csv -split "`n" | Select-Object -Skip 1) | ForEach-Object { ($_ -split "`t")[0] } | Where-Object { $_ })
    $added   = @($clean.Ticker | Where-Object { $prevTickers -notcontains $_ })
    $removed = @($prevTickers | Where-Object { $clean.Ticker -notcontains $_ })
    if ($added)   { Log ('Tickers added since last run: ' + ($added -join ', ')) }
    if ($removed) { Log ('Tickers removed since last run: ' + ($removed -join ', ')) }
  }
  foreach ($w in $warnings) { Log "WARN: $w" }

  # --- Write data.json (same tab-separated layout the page's parser already reads)
  $lines = @("Ticker`tDate`tTime`t`tCompany`t`t`tC/E")
  foreach ($row in $clean) { $lines += ($row.Ticker, $row.Date, $row.Time, $row.Timing, $row.Company, '', '', $row.CE) -join "`t" }
  $payload = [ordered]@{
    generatedAt = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    source      = "FactSet via '$SheetName' sheet of $(Split-Path -Leaf $WorkbookPath)"
    count       = $dated.Count
    tickers     = $clean.Count
    warnings    = @($warnings)
    csv         = ($lines -join "`n")
  }
  [IO.File]::WriteAllText($dataPath, ($payload | ConvertTo-Json -Depth 4), (New-Object Text.UTF8Encoding($false)))
  Log "Wrote $dataPath"

  # --- Publish --------------------------------------------------------------
  if ($NoPush) {
    Log 'NoPush set: skipping commit/push.'
    if ($EmailMode -eq 'Draft') { Send-TeamEmail 'Draft' }
    Log 'Done.'; return
  }
  $env:GIT_TERMINAL_PROMPT = '0'; $env:GCM_INTERACTIVE = 'never'
  Push-Location $RepoPath
  try {
    $branch = (git rev-parse --abbrev-ref HEAD).Trim()
    if ($branch -ne 'main') { throw "Repo is on branch '$branch', expected main" }
    git stash push --quiet -- data.json 2>&1 | Out-Null
    git pull --ff-only --quiet origin main 2>&1 | ForEach-Object { Log "git: $_" }
    if ($LASTEXITCODE -ne 0) { git stash pop --quiet 2>&1 | Out-Null; throw 'git pull --ff-only failed' }
    git checkout --quiet stash@{0} -- data.json 2>&1 | Out-Null; git stash drop --quiet 2>&1 | Out-Null
    git add data.json
    git commit --quiet -m ("Weekly earnings data refresh {0:yyyy-MM-dd}" -f (Get-Date)) -m "$($dated.Count) dated events from $($clean.Count) tickers. Automated by scripts/refresh-data.ps1." 2>&1 | ForEach-Object { Log "git: $_" }
    if ($LASTEXITCODE -ne 0) { Log 'Nothing to commit.' }
    git push --quiet origin main 2>&1 | ForEach-Object { Log "git: $_" }
    if ($LASTEXITCODE -ne 0) { throw 'git push failed' }
    Log ("Pushed {0}" -f (git rev-parse --short HEAD))
  } finally { Pop-Location }
  if ($EmailMode -ne 'None') {
    Log 'Waiting for GitHub Pages to serve the new data...'
    if (-not (Wait-SiteLive $payload.generatedAt)) { throw 'Pushed, but the live site did not show the new data within 10 minutes. Team email NOT sent.' }
    Log 'Live site confirmed.'
    Send-TeamEmail $EmailMode
  }
  if ($warnings.Count) { Notify 'Earnings calendar updated (with warnings)' ("{0} events published. {1} warning(s): see {2}" -f $dated.Count, $warnings.Count, $logFile) }
  Log 'Done.'
}
catch {
  Log "FAILED: $($_.Exception.Message)"
  Notify 'Earnings calendar refresh FAILED' ("$($_.Exception.Message) Site still shows last week's data. Log: $logFile")
  exit 1
}
finally {
  if ($xl) { try { $xl.Quit() } catch {}; try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl) } catch {} }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  if ($xlPid) { Start-Sleep -Seconds 2; Get-Process -Id $xlPid -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue }
  if (Test-Path $tempDir) { Remove-Item -Recurse -Force $tempDir -ErrorAction SilentlyContinue }
}
