# Example1 1: Get all shared mailbox storage quota and status
$mailbox_collection = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox | Sort-Object DisplayName
$mailbox_collection | .\Get-ExchangeMailboxSize.ps1 -Verbose | Export-Csv .\mailbox_size.csv -NoTypeInformation