
<#PSScriptInfo

.VERSION 0.2

.GUID 0899d54a-466b-4aa2-9234-94d3fbde54b6

.AUTHOR June Castillote

.COMPANYNAME

.COPYRIGHT

.TAGS

.LICENSEURI https://github.com/junecastillote/Exchange-Mailbox-Size-Export/blob/main/LICENSE

.PROJECTURI https://github.com/junecastillote/Exchange-Mailbox-Size-Export

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


#>

<#

.DESCRIPTION
 PowerShell script extract mailbox storage status in Exchange Online or Exchange Server

#>
[CmdletBinding()]
param (
    [Parameter(Mandatory, ValueFromPipeline)]
    $Mailbox,

    [Parameter()]
    [ValidateSet('KB', 'MB', 'GB', 'TB')]
    [string]
    $SizeUnitType = 'KB',

    [Parameter()]
    [bool]
    $FlattenResult = $true
)
begin {
    Function ConvertToUnit {
        param(
            [string] $SizeString,
            [string] $UnitType,
            [int] $RoundToDecimal = 2
        )
        $bytes = ([regex]::Match($SizeString, '\((\d{1,3}(,\d{3})*) bytes\)').Groups[1].Value).Replace(',', '')
        ([math]::Round($bytes / ("1$UnitType"), $RoundToDecimal) -as [double])
    }

    Function GetQuota {
        param(
            [PSObject] $quota,
            [string] $UnitType
        )
        if ($quota.IsUnlimited) {
            return 'Unlimited'
        }
        # return ConvertToUnit -SizeString $quota.Value -UnitType $UnitType
        return ConvertToUnit -SizeString $quota -UnitType $UnitType
    }
}
process {
    foreach ($item in $Mailbox) {
        if ($item.psobject.TypeNames -notcontains 'Deserialized.Microsoft.Exchange.Data.Directory.Management.Mailbox' `
                -and $item.psobject.TypeNames -notcontains 'Microsoft.Exchange.Data.Directory.Management.Mailbox' ) {
            Write-Error "The input is not a valid Exchange Mailbox object."
            continue
        }
        Write-Verbose "$($item.DisplayName)"

        $var_warning = ''
        $var_note = @("Size values are displayed in ($($SizeUnitType)).")
        try {
            $mailboxStatistics = Invoke-Command { Get-MailboxStatistics -Identity $item.Guid.ToString() } -WarningVariable var_warning
        }
        catch {
            Write-Error "Failed to retrieve mailbox statistics for $($item.DisplayName): $_"
            continue
        }

        $IssueWarningQuota = GetQuota -quota $item.IssueWarningQuota -UnitType $SizeUnitType
        $ProhibitSendQuota = GetQuota -quota $item.ProhibitSendQuota -UnitType $SizeUnitType
        $ProhibitSendReceiveQuota = GetQuota -quota $item.ProhibitSendReceiveQuota -UnitType $SizeUnitType

        $TotalItemSize = if ($var_warning -like "The user hasn't logged on to mailbox*") {
            $var_warning = "The TotalItemSize and PercentUsed values cannot be determined because the mailbox has no usage history or is empty."
            [double]0
        }
        else {
            ConvertToUnit -SizeString $mailboxStatistics.TotalItemSize -UnitType $SizeUnitType
        }

        $PercentUsed = if ($ProhibitSendReceiveQuota -ne 'Unlimited') {
            ([math]::Round(($TotalItemSize / $ProhibitSendReceiveQuota) * 100, 5) -as [double])
        }
        else {
            $var_note += 'PercentUsed value is zero (0) because the mailbox has no storage limit.'
            [double]0
        }

        $StorageLimitStatus = switch ($true) {
            { ($ProhibitSendReceiveQuota -ne 'Unlimited') -and ($TotalItemSize -ge $ProhibitSendReceiveQuota) } { 'Send and Receive Disabled' ; $var_note += "Mailbox storage is full. It cannot send or receive new items." }
            { ($ProhibitSendQuota -ne 'Unlimited') -and ($TotalItemSize -ge $ProhibitSendQuota) } { 'Send Disabled' ; $var_note += "Mailbox sending capability is disabled. It can still receive new items." }
            { ($IssueWarningQuota -ne 'Unlimited') -and ($TotalItemSize -ge $IssueWarningQuota) } { 'Warning' ; $var_note += "Mailbox storage is in warning status. It can still send and receive new items." }
            default { 'Normal' ; $var_note += "Mailbox size is below any quota." }
        }

        # Combine notes and warning messages
        $var_note = ($var_note + $var_warning)

        if ($FlattenResult) {
            # Convert notes from array to a flat numbered list string.
            # Example:
            #    [1] Message 1
            #    [2] Message 2
            for ($i = 0 ; $i -lt $var_note.Count; $i++) {
                $var_note[$i] = "[$($i+1)] $($var_note[$i])"
            }
            $var_note = $var_note -join "`n"
        }

        # return result object
        [PSCustomObject]@{
            DisplayName              = $item.DisplayName
            PrimaryEmailAddress      = $item.PrimarySmtpAddress.ToString()
            # ProxyAddresses           = $item.EmailAddresses -join ";"
            # ProxyAddresses           = $item.EmailAddresses.SmtpAddress -join ";"
            # ProxyEmailAddresses      = ($item.EmailAddresses | Where-Object { !$_.IsPrimaryAddress }).SmtpAddress -join ";"
            # ProxyEmailAddresses      = ($item.EmailAddresses | Where-Object { $_ -clike "smtp:*" })
            MailboxType              = $item.RecipientTypeDetails.ToString()
            IssueWarningQuota        = $IssueWarningQuota
            ProhibitSendQuota        = $ProhibitSendQuota
            ProhibitSendReceiveQuota = $ProhibitSendReceiveQuota
            TotalItemSize            = $TotalItemSize
            PercentUsed              = $PercentUsed
            StorageLimitStatus       = $StorageLimitStatus
            Notes                    = $var_note
        }
    }
}
end {

}

