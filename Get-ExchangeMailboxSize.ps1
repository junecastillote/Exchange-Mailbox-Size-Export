[CmdletBinding()]
param (
    [Parameter(Mandatory, ValueFromPipeline)]
    $Mailbox
)
begin {
    Function MailboxSizeStringToNumber {
        param(
            [Parameter(Mandatory)]
            [string]
            $SizeString,

            [Parameter()]
            [ValidateSet('KB', 'MB', 'GB', 'TB')]
            [string]
            $UnitType = 'GB',

            [Parameter()]
            [int]
            $RoundToDecimal = 2
        )
        $regEx_Pattern = '\((\d{1,3}(,\d{3})*) bytes\)'
        $bytesString = ([regex]::Match($SizeString, $regEx_Pattern).Groups[1].Value).Replace(',', '')
        [System.Math]::Round(([int64]$bytesString / "1$($UnitType)"), $RoundToDecimal)
    }
}
process {
    foreach ($item in $Mailbox) {
        if ($item.psobject.TypeNames -notcontains 'Deserialized.Microsoft.Exchange.Data.Directory.Management.Mailbox') {
            Write-Error "The input is not a valid Exchange Mailbox object."
            continue
        }

        try {
            $mailbox_statistics = Get-MailboxStatistics -Identity $item.Guid
        }
        catch {
            Write-Error "Failed to retrieve mailbox statistics for $($item.DisplayName): $_"
            continue
        }

        $IssueWarningQuota = MailboxSizeStringToNumber -SizeString $item.IssueWarningQuota
        $ProhibitSendQuota = MailboxSizeStringToNumber -SizeString $item.ProhibitSendQuota
        $ProhibitSendReceiveQuota = MailboxSizeStringToNumber -SizeString $item.ProhibitSendReceiveQuota
        $TotalItemSize = MailboxSizeStringToNumber -SizeString $mailbox_statistics.TotalItemSize
        $PercentUsed = [System.Math]::Round((($TotalItemSize / $ProhibitSendReceiveQuota) * 100), 2)

        # Determine the storage limit status using switch
        $StorageLimitStatus = switch ($true) {
            { $TotalItemSize -ge $ProhibitSendReceiveQuota } { 'Send and Receive Disabled'; break }
            { $TotalItemSize -ge $ProhibitSendQuota } { 'Send Disabled'; break }
            { $TotalItemSize -ge $IssueWarningQuota } { 'Warning'; break }
            default { 'Normal' }
        }

        [PSCustomObject]@{
            DisplayName              = $item.DisplayName
            PrimarySmtpAddress       = $item.PrimarySmtpAddress
            IssueWarningQuota        = $IssueWarningQuota
            ProhibitSendQuota        = $ProhibitSendQuota
            ProhibitSendReceiveQuota = $ProhibitSendReceiveQuota
            TotalItemSize            = $TotalItemSize
            PercentUsed              = $PercentUsed
            StorageLimitStatus       = $StorageLimitStatus
        }
        Write-Verbose "Processed mailbox: $($item.DisplayName)"
    }
}
end {

}
