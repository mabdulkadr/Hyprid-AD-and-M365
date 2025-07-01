<#
.SYNOPSIS
    Synchronize on-prem Active Directory groups with Entra ID or Exchange Online distribution lists.

.DESCRIPTION
    - Supports syncing Security, Microsoft 365, and Distribution groups.
    - Uses certificate-based app authentication for Microsoft Graph and Exchange Online.
    - Retrieves all members (no paging issues).
    - Efficient comparison using HashSets.
    - Logs structured, readable results with color-coded output.

.NOTES
    Author  : Mohammed Omar
    Version : 2.3
    Date    : 2025-05-22

.REQUIREMENTS
    - Modules: ExchangeOnlineManagement, Microsoft.Graph
    - AD RSAT Tools
    - App Registration with permissions:
        • Microsoft Graph: Group.Read.All, GroupMember.ReadWrite.All, User.Read.All
        • Exchange Online: Application-level access
#>

# ================== Variables ==================

#Log File Name
$LogFileName = "xxxxxxxxxxxxx"

# Mapping: On-prem AD group mail to corresponding Entra or Exchange group ID
$GroupMappings = @{
    "AD-Group@qu.edu.sa" = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"

}


# ================== Configuration ==================
$TenantId              = "c2b04da6-8487-41cc-8803-90321048a772"
$AppId                 = "6c70c0c3-e3a6-489c-973e-51e8138540f9"
$CertificateThumbprint = "93B8050D00970041F1D97C7A58DFE7358252ACAC"
$Organization          = "qu.edu.sa"

# ================== LOGGING ==================
$BasePath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$LogFolder = Join-Path $BasePath "Logs"
if (!(Test-Path $LogFolder)) { New-Item -Path $LogFolder -ItemType Directory -Force }
$LogFile = Join-Path $LogFolder "Sync-Groups-$LogFileName-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "ERROR", "WARNING")]
        [string]$Level = "INFO"
    )
    $Color = switch ($Level) {
        "INFO"     { "Cyan" }
        "SUCCESS"  { "Green" }
        "WARNING"  { "Yellow" }
        "ERROR"    { "Red" }
    }

    if ($Level -in @("ERROR", "WARNING")) {
        $entry = "[$Level] $Message"
    } else {
        $entry = "$Message"
    }

    Write-Host $entry -ForegroundColor $Color
    Add-Content -Path $LogFile -Value $entry -Encoding UTF8
}


function Write-Section {
    param ([string]$Title)
    $Line = "=" * 80
    Write-Log "`n$Line" "INFO"
    Write-Log "$Title" "INFO"
    Write-Log "$Line" "INFO"
}

# ================== CONNECT ==================
Write-Log "🔄 Connecting to Microsoft Graph..." "INFO"
Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
Write-Log "✅ Connected to Microsoft Graph." "SUCCESS"

Write-Log "🔄 Connecting to Exchange Online..." "INFO"
Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false
Write-Log "✅ Connected to Exchange Online." "SUCCESS"

# ================== MAIN SYNC LOOP ==================
foreach ($mapping in $GroupMappings.GetEnumerator()) {
    $ADGroupName   = $mapping.Key
    $CloudGroupId  = $mapping.Value

    Write-Section "Group: $ADGroupName  -->  $CloudGroupId"

    $EnabledUPNs   = [System.Collections.Generic.HashSet[string]]::new()
    $DisabledUPNs  = @()
    $AddedUsers    = @()
    $RemovedUsers  = @()
    $CloudUPNs     = [System.Collections.Generic.HashSet[string]]::new()

    # === Get AD Members ===
    try {
        try {
            $ADGroup = Get-ADGroup -Identity $ADGroupName -ErrorAction Stop
        } catch {
            $ADGroup = Get-ADGroup -Filter "mail -eq '$ADGroupName'" -ErrorAction Stop
        }

        $Members = (Get-ADGroup -Identity $ADGroup -Properties Member).Member
        foreach ($dn in $Members) {
            try {
                $user = Get-ADUser -Identity $dn -Properties UserPrincipalName, Enabled
                if ($user.Enabled -and $user.UserPrincipalName) {
                    $null = $EnabledUPNs.Add($user.UserPrincipalName.ToLower())
                } elseif ($user.UserPrincipalName) {
                    $DisabledUPNs += $user.UserPrincipalName.ToLower()
                }
            } catch {
                Write-Log "⚠️ Skipped invalid member: $dn" "WARNING"
            }
        }

        Write-Log "📥 Retrieved $($EnabledUPNs.Count + $DisabledUPNs.Count) users from AD ($($EnabledUPNs.Count) enabled, $($DisabledUPNs.Count) disabled)." "SUCCESS"
    } catch {
        Write-Log "❌ Failed to retrieve AD group members: $($_.Exception.Message)" "ERROR"
        continue
    }

    # ================== Get Cloud Members (via Graph) ==================
    try {
        $group = Get-MgGroup -GroupId $CloudGroupId -ErrorAction Stop

        $uri = "https://graph.microsoft.com/v1.0/groups/$CloudGroupId/members?$top=999"
        do {
            $result = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            foreach ($member in $result.value) {
                if ($member.'@odata.type' -eq "#microsoft.graph.user" -and $member.userPrincipalName) {
                    $null = $CloudUPNs.Add($member.userPrincipalName.ToLower())
                }
            }
            $uri = $result.'@odata.nextLink'
        } while ($uri)

        Write-Log "📤 Retrieved $($CloudUPNs.Count) users from cloud group." "SUCCESS"
    } catch {
        Write-Log "❌ Error retrieving cloud group members: $($_.Exception.Message)" "ERROR"
        continue
    }

    # ================== Compare & Sync ==================
    $toAdd = [System.Collections.Generic.HashSet[string]]::new($EnabledUPNs)
    $toAdd.ExceptWith($CloudUPNs)

    $toRemove = [System.Collections.Generic.HashSet[string]]::new($CloudUPNs)
    $toRemove.ExceptWith($EnabledUPNs)

     foreach ($upn in $toAdd) {
    try {
        if ($group.GroupTypes.Count -eq 0 -and $group.MailEnabled) {
            # Exchange Distribution Group
            Add-DistributionGroupMember -Identity $group.id -Member $upn -ErrorAction Stop
            $AddedUsers += $upn
        }
        else {
            # Microsoft 365 or Security Group
            $user = Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction Stop
            New-MgGroupMember -GroupId $CloudGroupId -DirectoryObjectId $user.Id -ErrorAction Stop
            $AddedUsers += $upn
        }
    } catch {
        Write-Log "❌ Failed to add $upn : $($_.Exception.Message)" "ERROR"
    }
}

    foreach ($upn in $toRemove) {
    try {
        if ($group.GroupTypes.Count -eq 0 -and $group.MailEnabled) {
            # Exchange Distribution Group
            Remove-DistributionGroupMember -Identity $group.id -Member $upn -Confirm:$false -ErrorAction Stop
            $RemovedUsers += $upn
        }
        else {
            # Microsoft 365 or Security Group
            $user = Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction Stop
            Remove-MgGroupMemberByRef -GroupId $CloudGroupId -DirectoryObjectId $user.Id -ErrorAction Stop
            $RemovedUsers += $upn
        }
    } catch {
        Write-Log "❌ Failed to remove $upn : $($_.Exception.Message)" "ERROR"
    }
}

    # ================== Summary ==================
    if ($AddedUsers.Count -gt 0) {
        Write-Log "✅ Added users:" "SUCCESS"
        $AddedUsers | ForEach-Object { Write-Log "   + $_" "INFO" }
    } else {
        Write-Log "ℹ️ No users added." "INFO"
    }

    if ($RemovedUsers.Count -gt 0) {
        Write-Log "⚠️ Removed users:" "WARNING"
        $RemovedUsers | ForEach-Object { Write-Log "   - $_" "INFO" }
    } else {
        Write-Log "ℹ️ No users removed." "INFO"
    }

    Write-Log ""
    Write-Log "---------------------------------------------" "INFO"
    Write-Log "📋 Group Summary:" "INFO"
    Write-Log " - AD Total: $($EnabledUPNs.Count + $DisabledUPNs.Count)" "INFO"
    Write-Log " - AD Enabled: $($EnabledUPNs.Count)" "INFO"
    Write-Log " - AD Disabled: $($DisabledUPNs.Count)" "INFO"
    Write-Log " - Cloud Members: $($CloudUPNs.Count)" "INFO"
    Write-Log " - Added: $($AddedUsers.Count), Removed: $($RemovedUsers.Count)" "INFO"
    Write-Log "---------------------------------------------" "INFO"
}

# ================== DISCONNECT ==================
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Disconnect-MgGraph -ErrorAction SilentlyContinue
Write-Log "`n✅ All group syncs completed. Log saved to:`n$LogFile" "SUCCESS"
