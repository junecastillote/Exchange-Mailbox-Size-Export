[CmdletBinding()]
param (
    [Parameter(Mandatory, ValueFromPipeline)]
    $Mailbox,

    [Parameter()]
    [ValidateSet('KB', 'MB', 'GB', 'TB')]
    [string]
    $SizeUnitType = 'KB'
)
begin {
    Function ConvertToUnit {
        param(
            [string] $SizeString,
            [string] $UnitType,
            [int] $RoundToDecimal = 2
        )
        $bytes = ([regex]::Match($SizeString, '\((\d{1,3}(,\d{3})*) bytes\)').Groups[1].Value).Replace(',', '')
        [math]::Round($bytes / ("1$UnitType"), $RoundToDecimal)
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

        try {
            $mailboxStatistics = Get-MailboxStatistics -Identity $item.Guid.ToString() -WarningVariable warning -WarningAction SilentlyContinue
        }
        catch {
            Write-Error "Failed to retrieve mailbox statistics for $($item.DisplayName): $_"
            continue
        }

        $IssueWarningQuota = GetQuota -quota $item.IssueWarningQuota -UnitType $SizeUnitType
        $ProhibitSendQuota = GetQuota -quota $item.ProhibitSendQuota -UnitType $SizeUnitType
        $ProhibitSendReceiveQuota = GetQuota -quota $item.ProhibitSendReceiveQuota -UnitType $SizeUnitType

        $TotalItemSize = if ($warning[0] -like "The user hasn't logged on to mailbox*") {
            0
        } else {
            ConvertToUnit -SizeString $mailboxStatistics.TotalItemSize -UnitType $SizeUnitType
            # ConvertToUnit -SizeString $mailboxStatistics.TotalItemSize.Value -UnitType $SizeUnitType
        }

        $PercentUsed = if ($ProhibitSendReceiveQuota -ne 'Unlimited') {
            [math]::Round(($TotalItemSize / $ProhibitSendReceiveQuota) * 100, 2)
        } else {
            0
        }

        $StorageLimitStatus = switch ($true) {
            { ($ProhibitSendReceiveQuota -ne 'Unlimited') -and ($TotalItemSize -ge $ProhibitSendReceiveQuota) } { 'Send and Receive Disabled' }
            { ($ProhibitSendQuota -ne 'Unlimited') -and ($TotalItemSize -ge $ProhibitSendQuota) } { 'Send Disabled' }
            { ($IssueWarningQuota -ne 'Unlimited') -and ($TotalItemSize -ge $IssueWarningQuota) } { 'Warning' }
            default { 'Normal' }
        }

        [PSCustomObject]@{
            DisplayName              = $item.DisplayName
            PrimarySmtpAddress       = $item.PrimarySmtpAddress#.ToString()
            IssueWarningQuota        = $IssueWarningQuota
            ProhibitSendQuota        = $ProhibitSendQuota
            ProhibitSendReceiveQuota = $ProhibitSendReceiveQuota
            TotalItemSize            = $TotalItemSize
            PercentUsed              = $PercentUsed
            StorageLimitStatus       = $StorageLimitStatus
            Notes                    = $warning -join "`n"
        }

        Write-Verbose "Processed mailbox: $($item.DisplayName)"
    }
}
end {

}
