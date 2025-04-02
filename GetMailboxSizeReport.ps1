<#
=============================================================================================
Name:           Microsoft 365 Mailbox Size Report
Version:        2.4
Fixed Issues:   Empty filter handling, proper mailbox retrieval
============================================================================================
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory = $false)]
    [switch]$MFA,
    
    [Parameter(Mandatory = $false)]
    [switch]$DeviceAuth,
    
    [Parameter(Mandatory = $false)]
    [switch]$SharedMBOnly,
    
    [Parameter(Mandatory = $false)]
    [switch]$UserMBOnly,
    
    [Parameter(Mandatory = $false)]
    [string]$MBNamesFile,
    
    [Parameter(Mandatory = $false)]
    [string]$UserName,
    
    [Parameter(Mandatory = $false)]
    [string]$Password,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "."
)

Begin {
    $ErrorActionPreference = "Stop"
    $ScriptStartTime = Get-Date
    $ExportCSV = Join-Path $OutputPath "MailboxSizeReport_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').csv"
    $Results = @()
    $MailboxCount = 0

    function Connect-EXO {
        if ($DeviceAuth -or ($MFA -and -not [Environment]::UserInteractive)) {
            Write-Host "Using device code authentication..." -ForegroundColor Cyan
            Connect-ExchangeOnline -Device -ShowBanner:$false -ErrorAction Stop
        }
        elseif ($MFA) {
            Write-Host "Using interactive authentication..." -ForegroundColor Cyan
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
        elseif ($UserName -and $Password) {
            $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
            $Credential = New-Object System.Management.Automation.PSCredential($UserName, $SecuredPassword)
            Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false -ErrorAction Stop
        }
        else {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
    }

    function Cleanup {
        try {
            Write-Host "Disconnecting session..." -ForegroundColor Cyan
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        }
        catch {
            Write-Warning "Disconnection error: $_"
        }
    }

    function ConvertTo-Bytes {
        param([string]$SizeString)
        if ([string]::IsNullOrEmpty($SizeString)) { return 0 }
        
        $multiplier = 1
        if ($SizeString -match "KB") { $multiplier = 1KB }
        elseif ($SizeString -match "MB") { $multiplier = 1MB }
        elseif ($SizeString -match "GB") { $multiplier = 1GB }
        elseif ($SizeString -match "TB") { $multiplier = 1TB }
        
        $value = [double]($SizeString -replace "[^\d\.]")
        return [math]::Round($value * $multiplier)
    }
}

Process {
    try {
        # Install module if needed
        if (-not (Get-Module ExchangeOnlineManagement -ListAvailable)) {
            Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
            Import-Module ExchangeOnlineManagement -Force
        }

        # Connect to Exchange Online
        Connect-EXO
        Write-Host "Connected successfully" -ForegroundColor Green

        # Get mailboxes based on input parameters
        if ($MBNamesFile -and (Test-Path $MBNamesFile)) {
            Write-Host "Processing mailboxes from input file..." -ForegroundColor Cyan
            $Mailboxes = Import-Csv -Header "Identity" $MBNamesFile | ForEach-Object { $_.Identity }
        }
        else {
            Write-Host "Retrieving mailboxes from Exchange Online..." -ForegroundColor Cyan
            
            # Build filter based on parameters
            $FilterParams = @()
            if ($SharedMBOnly) { $FilterParams += "RecipientTypeDetails -eq 'SharedMailbox'" }
            if ($UserMBOnly) { $FilterParams += "RecipientTypeDetails -eq 'UserMailbox'" }
            
            $Filter = $FilterParams -join " -or "
            if (-not $Filter) { $Filter = $null }

            if ($Filter) {
                Write-Host "Using filter: $Filter" -ForegroundColor DarkGray
                $Mailboxes = Get-Mailbox -ResultSize Unlimited -Filter $Filter | Select-Object -ExpandProperty UserPrincipalName
            }
            else {
                $Mailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object -ExpandProperty UserPrincipalName
            }
        }

        $TotalMailboxes = $Mailboxes.Count
        if ($TotalMailboxes -eq 0) {
            Write-Host "No mailboxes found matching the specified criteria" -ForegroundColor Yellow
            return
        }

        Write-Host "Found $TotalMailboxes mailboxes to process" -ForegroundColor Cyan

        # Process each mailbox
        foreach ($UPN in $Mailboxes) {
            $MailboxCount++
            Write-Progress -Activity "Processing mailboxes" -Status "$MailboxCount of $TotalMailboxes" -CurrentOperation $UPN -PercentComplete (($MailboxCount / $TotalMailboxes) * 100)

            try {
                $MBDetails = Get-Mailbox -Identity $UPN -ErrorAction Stop
                $Stats = Get-MailboxStatistics -Identity $UPN -ErrorAction Stop

                $Result = [PSCustomObject]@{
                    'Display Name' = $MBDetails.DisplayName
                    'User Principal Name' = $MBDetails.UserPrincipalName
                    'Mailbox Type' = $MBDetails.RecipientTypeDetails
                    'Primary SMTP Address' = $MBDetails.PrimarySmtpAddress
                    'Item Count' = $Stats.ItemCount
                    'Total Size' = $Stats.TotalItemSize.Value -replace "\(.*"
                    'Total Size (Bytes)' = ConvertTo-Bytes $Stats.TotalItemSize.Value
                    'Deleted Item Count' = $Stats.DeletedItemCount
                    'Deleted Item Size' = $Stats.TotalDeletedItemSize.Value -replace "\(.*"
                    'Deleted Item Size (Bytes)' = ConvertTo-Bytes $Stats.TotalDeletedItemSize.Value
                    'Archive Status' = if ($MBDetails.ArchiveStatus -eq "Active") { "Active" } else { "Disabled" }
                    'Issue Warning Quota' = if ($MBDetails.IssueWarningQuota) { $MBDetails.IssueWarningQuota -replace "\(.*" } else { "Not Set" }
                    'Prohibit Send Quota' = if ($MBDetails.ProhibitSendQuota) { $MBDetails.ProhibitSendQuota -replace "\(.*" } else { "Not Set" }
                    'Prohibit Send Receive Quota' = if ($MBDetails.ProhibitSendReceiveQuota) { $MBDetails.ProhibitSendReceiveQuota -replace "\(.*" } else { "Not Set" }
                    'Last Logon Time' = if ($Stats.LastLogonTime) { $Stats.LastLogonTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
                }

                $Results += $Result
            }
            catch {
                Write-Warning "Error processing mailbox $UPN : $_"
                continue
            }
        }

        # Export results
        if ($Results.Count -gt 0) {
            $Results | Export-Csv -Path $ExportCSV -NoTypeInformation -Encoding UTF8
            Write-Host "`nReport generated successfully!" -ForegroundColor Green
            Write-Host "Processed $MailboxCount mailboxes" -ForegroundColor Cyan
            Write-Host "Output file: $ExportCSV" -ForegroundColor Yellow

            # Offer to open the file
            $OpenFile = Read-Host "Do you want to open the output file now? [Y] Yes [N] No"
            if ($OpenFile -match "[yY]") {
                Start-Process $ExportCSV
            }
        }
        else {
            Write-Host "No mailbox data was exported." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "`nERROR: $_" -ForegroundColor Red
        Write-Host "Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
        if ($_.Exception -like "*window handle*") {
            Write-Host "`nSOLUTION: Rerun with -DeviceAuth parameter" -ForegroundColor Yellow
            Write-Host "Example: .\GetMailboxSizeReport.ps1 -MFA -DeviceAuth"
        }
    }
    finally {
        Cleanup
        Write-Host "Execution time: $((Get-Date).Subtract($ScriptStartTime).TotalMinutes.ToString('0.0')) minutes" -ForegroundColor Cyan
    }
}