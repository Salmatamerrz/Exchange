<#
.SYNOPSIS
  Bulk deletes emails from Exchange Server based on From, To, and Subject.
 
.DESCRIPTION
  Reads a CSV file containing sender (From), recipient mailbox (To) and Subject.
  Executes Search-Mailbox to locate and delete the matching messages.
  Uses the CSV path C:\Users\msx.support\Desktop\Delete.csv by default but can be overridden.
 
.PARAMETER CsvPath
  Path to the CSV file (defaults to C:\Users\msx.support\Desktop\Delete.csv).
 
.PARAMETER LogPath
  Optional path to save the transcript/log file. If not provided, a file is created
  in the same directory as the script (or the current working directory if the script path is unavailable).
 
.EXAMPLE
  .\Delete‑Emails‑Exchange.ps1
  .\Delete‑Emails‑Exchange.ps1 -CsvPath "D:\Other\emails.csv"
#>
 
[CmdletBinding()]
param (
    [ValidateNotNullOrEmpty()]
    [string]$CsvPath = 'C:\Users\msx.support\Desktop\delete(in).csv',
 
    [string]$LogPath
)
 
# Determine a default log path if one isn't provided.
if (-not $LogPath) {
    # $PSScriptRoot is defined when the script is run from a file.
    $logDir = $PSScriptRoot
    if (-not $logDir) {
        # Fallback to the current working directory if $PSScriptRoot is null.
        $logDir = (Get-Location).Path
    }
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $LogPath = Join-Path -Path $logDir -ChildPath ("DeleteEmailsLog_{0}.txt" -f $timestamp)
}
 
# Attempt to load Exchange snap-ins when outside Exchange Management Shell
$exchangeSnapins = @(
    'Microsoft.Exchange.Management.PowerShell.SnapIn',
    'Microsoft.Exchange.Management.PowerShell.E2010',
    'Microsoft.Exchange.Management.PowerShell.E2013',
    'Microsoft.Exchange.Management.PowerShell.E2016',
    'Microsoft.Exchange.Management.PowerShell.E2019'
)
foreach ($snapin in $exchangeSnapins) {
    if (-not (Get-PSSnapin -Name $snapin -ErrorAction SilentlyContinue)) {
        try {
            Add-PSSnapin $snapin -ErrorAction Stop
            Write-Verbose "Loaded Exchange snap-in: $snapin"
            break
        } catch {
            # ignore and try the next snap-in
        }
    }
}
 
# Start transcript logging
Start-Transcript -Path $LogPath -Append
 
try {
    # Validate CSV file exists
    if (-not (Test-Path $CsvPath)) {
        throw "CSV file not found: $CsvPath"
    }
 
    # Load CSV data
    $items = Import-Csv -Path $CsvPath
    if (-not $items) {
        throw "CSV file is empty or missing required columns (From, To, Subject)."
    }
 
    foreach ($entry in $items) {
        $from    = $entry.From
        $to      = $entry.To
        $subject = $entry.Subject
 
        if ([string]::IsNullOrWhiteSpace($from) -or [string]::IsNullOrWhiteSpace($to) -or [string]::IsNullOrWhiteSpace($subject)) {
            Write-Warning "Skipping invalid row: $($entry | ConvertTo-Csv -NoTypeInformation)"
            continue
        }
 
        Write-Host "Processing mailbox: $to`nFrom: $from`nSubject: $subject" -ForegroundColor Cyan
 
        try {
            # Build search query
            $searchQuery = "From:`"$from`" AND Subject:`"$subject`""
 
            # Run Search-Mailbox to delete content
            Search-Mailbox -Identity $to `
                           -SearchQuery $searchQuery `
                           -DeleteContent `
                           -Force `
                           -ErrorAction Stop
 
            Write-Host "✔ Successfully deleted matching emails from $to." -ForegroundColor Green
        } catch {
            Write-Error "Failed to delete emails for $to. Error: $_"
        }
    }
} catch {
    Write-Error "Fatal script error: $_"
} finally {
    Stop-Transcript
    Write-Host "`nTranscript saved to: $LogPath" -ForegroundColor Yellow
}