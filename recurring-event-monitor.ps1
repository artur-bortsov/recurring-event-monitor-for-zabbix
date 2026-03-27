<#
.SYNOPSIS
    Recurring Windows Event Log Error Monitor with Zabbix trap notifications.

.DESCRIPTION
    Scans the Windows event log for Error (Level 2) and Warning (Level 3) events.
    Whenever a specific Event ID / Source combination occurs more than CompareCount
    times within CompareInterval minutes, a trap is sent to the Zabbix server via
    zabbix_sender.

    A flood-protection fuse (FloodFuse) prevents the same pair from generating
    more than one alert per FloodFuse-minute window, avoiding alert storms.

    The script maintains two cache files in %SystemRoot%\Temp between runs:
      Config.cache  – timestamp marking the start of the next query window
      Loghash.xml   – serialised hash table of per-event sliding windows and fuses

    Designed to be triggered by the Zabbix agent on a regular schedule
    (e.g., every hour) using a system.run item with the "nowait" option.

.NOTES
    Version : 1.2
    Author  : Bortsov A.S.
    License : GNU GPL v3

.LINK
    https://github.com/artur-bortsov/recurring-event-monitor-for-zabbix
#>

# ---------------------------------------------------------------------------
# TimeTable – per-event-key state object.
# Holds the sliding window of event timestamps (TimeArray) and the fuse
# timestamp that suppresses repeated alerts (FuseTime).
# ---------------------------------------------------------------------------
class TimeTable {
    [ValidateNotNullOrEmpty()][System.DateTime]           $FuseTime
    [ValidateNotNull()]       [System.Collections.ArrayList] $TimeArray

    TimeTable([System.Collections.ArrayList]$TArray, [System.DateTime]$FTime) {
        $this.TimeArray = $TArray
        $this.FuseTime  = $FTime
    }
}

# ---------------------------------------------------------------------------
# Configuration – loaded from config.psd1 next to the script file.
# Built-in defaults are used if the file is absent.
# ---------------------------------------------------------------------------
$configFile = Join-Path $PSScriptRoot "config.psd1"
if (Test-Path $configFile) {
    $cfg              = Import-PowerShellDataFile -Path $configFile
    $CompareInterval  = [int]$cfg.CompareInterval
    $CompareCount     = [int]$cfg.CompareCount
    $FloodFuse        = [int]$cfg.FloodFuse
    $ZabbixSenderPath = $cfg.ZabbixSenderPath
    $ZabbixConfigPath = $cfg.ZabbixConfigPath
    $MonitoredLogs    = $cfg.EventLogs
} else {
    $CompareInterval  = 10
    $CompareCount     = 5
    $FloodFuse        = 1440
    $ZabbixSenderPath = "C:\Program Files\ZabbixAgent\zabbix_sender.exe"
    $ZabbixConfigPath = "C:\Program Files\ZabbixAgent\zabbix_agentd.win.conf"
    $MonitoredLogs    = @("System", "Application")
}

# ---------------------------------------------------------------------------
# Exceptions – loaded from exceptions.conf next to the script file.
# Format: EventID:SourceName [# optional comment]
# Lines that start with # or are blank are ignored.
# Inline # comments are stripped before the entry is stored.
# Matching is case-insensitive.
# ---------------------------------------------------------------------------
$exceptionsFile = Join-Path $PSScriptRoot "exceptions.conf"
$Exceptions = @()
if (Test-Path $exceptionsFile) {
    Get-Content $exceptionsFile | ForEach-Object {
        # Strip inline comments and surrounding whitespace.
        $line = ($_ -replace '#.*$', '').Trim()
        # Only accept lines that look like a valid "digits:something" entry.
        if ($line -match '^\d+:.+') {
            $Exceptions += $line.ToLower()
        }
    }
}

# ---------------------------------------------------------------------------
# Current date – captured once at script start and reused throughout the run.
# ---------------------------------------------------------------------------
$CurrentDate = Get-Date

# ---------------------------------------------------------------------------
# Persistent state cache paths.
# Both files live in %SystemRoot%\Temp which is writable by the SYSTEM account
# that runs the Zabbix agent service.
# ---------------------------------------------------------------------------
$cacheDir      = "$env:SystemRoot\Temp"
$checkDateFile = Join-Path $cacheDir "Config.cache"
$logHashFile   = Join-Path $cacheDir "Loghash.xml"

# Load the start-of-window timestamp written by the previous run.
# Defaults to 24 hours ago on the first run or if the cache is corrupt.
[System.DateTime]$checkDate = $null
if (Test-Path $checkDateFile) {
    try {
        [System.DateTime]$checkDate = Get-Date (Get-Content $checkDateFile)
    } catch { }
}
if (-not $checkDate) {
    $checkDate = $CurrentDate.AddDays(-1)
}

# Load the persisted event hash from the previous run and purge any entries
# that have since been added to the exceptions list.
$logHash = $null
if (Test-Path $logHashFile) {
    $logHash = Import-Clixml $logHashFile
    foreach ($exc in $Exceptions) {
        if ($logHash.Keys -contains $exc) {
            $logHash.Remove($exc)
        }
    }
}
if (-not $logHash) {
    $logHash = @{}
}

# ---------------------------------------------------------------------------
# Get-EventLogs – returns all Error and Warning events from the monitored
# channels within the configured time window.
# Level 2 = Error, Level 3 = Warning.
# Milliseconds are zeroed out to avoid off-by-one boundary issues with the
# Windows event log query engine.
# ---------------------------------------------------------------------------
function Get-EventLogs {
    param (
        [string[]]$LogName   = $MonitoredLogs,
        [datetime]$StartTime = $checkDate,
        [datetime]$EndTime   = $CurrentDate
    )

    $start = $StartTime.AddMilliseconds(-$StartTime.Millisecond)
    $end   = $EndTime.AddMilliseconds(-$EndTime.Millisecond)

    Get-WinEvent -FilterHashtable @{
        LogName   = $LogName
        Level     = 2, 3
        StartTime = $start
        EndTime   = $end
    } -ErrorAction SilentlyContinue | Select-Object *
}

# ---------------------------------------------------------------------------
# Get-EventLevelName – converts the numeric event level to a readable label.
# ---------------------------------------------------------------------------
function Get-EventLevelName {
    param ([int]$LevelID)
    if ($LevelID -eq 2) { "Error" } else { "Warning" }
}

# Collect events for the time window since the last run.
$EventLogs = Get-EventLogs

# Reverse to chronological order (oldest first) so the sliding window
# correctly sees events in the order they occurred.
if ($EventLogs.Count -gt 1) {
    [array]::Reverse($EventLogs)
}

# Filter out every event whose "EventID:SourceName" key matches an exceptions entry.
$EventLogsFiltered = $EventLogs | Where-Object {
    -not $Exceptions.Contains(("$($_.Id):$($_.ProviderName)").ToLower())
}

# ---------------------------------------------------------------------------
# Main detection loop.
# For each filtered event, maintain a per-key sliding window of timestamps.
# When the window contains CompareCount or more occurrences within
# CompareInterval minutes AND the flood-protection fuse has expired,
# send a Zabbix trap and reset the fuse.
# ---------------------------------------------------------------------------
foreach ($event in $EventLogsFiltered) {

    # The hash key is "EventID:SourceName" (case-sensitive storage, but the
    # exceptions list is matched case-insensitively above).
    $wideID = "$($event.Id):$($event.ProviderName)"

    # Create a new empty TimeTable entry for previously-unseen event keys.
    if ($logHash.Keys -notcontains $wideID) {
        $logHash[$wideID] = [TimeTable]::new(
            [System.Collections.ArrayList]@(),
            [datetime]"01.01.2001 00:00:00"
        )
    }

    $entry = $logHash[$wideID]

    # Add the current event's timestamp to the sliding window.
    $entry.TimeArray.Add($event.TimeCreated) | Out-Null

    # Drop timestamps that are older than CompareInterval minutes relative
    # to the current event's time (i.e., outside the sliding window).
    while ($entry.TimeArray.Count -gt 0 -and
           ($event.TimeCreated - $entry.TimeArray[0]).TotalMinutes -gt $CompareInterval) {
        $entry.TimeArray.RemoveAt(0)
    }

    # Trigger an alert if the event count meets the threshold and the fuse
    # cooldown period has elapsed since the last alert for this key.
    if ($event.TimeCreated -gt $entry.FuseTime -and
        $entry.TimeArray.Count -ge $CompareCount) {

        # Remove single and double quotes to avoid breaking the zabbix_sender
        # command-line argument parsing.
        $cleanMessage = $event.Message -replace '"', '' -replace "'", ''

        # Re-encode the message bytes: some legacy event sources (especially
        # older Cyrillic-locale applications) embed Windows-1251-encoded text
        # in the event record. PowerShell reads it as UTF-8 Unicode code points,
        # which appear garbled. Converting through UTF-8 → Windows-1251 restores
        # the original characters for those sources while leaving properly-encoded
        # UTF-8 messages mostly intact.
        [string]$message = [Text.Encoding]::GetEncoding("windows-1251").GetString(
            [Text.Encoding]::GetEncoding("UTF-8").GetBytes($cleanMessage)
        )

        # Build the human-readable alert payload that will appear in Zabbix.
        $payload = "Journal: $($event.LogName)`n" +
                   "Level: $(Get-EventLevelName -LevelID $event.Level)`n" +
                   "Source: $($event.ProviderName)`n" +
                   "ID: $($event.Id)`n" +
                   "Time: $($event.TimeCreated.ToString('dd.MM.yyyy HH:mm:ss'))`n`n" +
                   "Message: $message"

        # Send the trap to the Zabbix server using zabbix_sender.
        # -c  : agent config file; zabbix_sender reads Server= from it automatically
        # -s  : host name as registered in Zabbix (uppercase computer name)
        # -k  : Zabbix item key that will receive the value
        # -o  : the value to send
        & $ZabbixSenderPath `
            -c $ZabbixConfigPath `
            -s ($env:COMPUTERNAME).ToUpper() `
            -k repeat.monitor `
            -o $payload

        # Arm the fuse: suppress further alerts for this key for FloodFuse minutes.
        $entry.FuseTime = $event.TimeCreated.AddMinutes($FloodFuse)
    }
}

# ---------------------------------------------------------------------------
# Persist state for the next run.
# ---------------------------------------------------------------------------

# Save the current timestamp; the next run will query events starting here.
$CurrentDate | Out-File $checkDateFile

# Serialise the full event hash table to XML so the sliding windows and fuse
# times survive between script invocations.
$logHash | Export-Clixml $logHashFile
