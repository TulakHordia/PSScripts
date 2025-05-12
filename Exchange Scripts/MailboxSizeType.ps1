$Mailboxes = Get-Mailbox -ResultSize Unlimited

$MailboxData = $Mailboxes | ForEach-Object {
    $Stats = Get-MailboxStatistics $_.PrimarySmtpAddress
    $ArchiveStats = Get-MailboxStatistics $_.PrimarySmtpAddress -Archive
    [PSCustomObject]@{
        DisplayName      = $_.DisplayName
        PrimarySMTP      = $_.PrimarySmtpAddress
        MailboxType      = $_.RecipientTypeDetails
        TotalItemSize    = $Stats.TotalItemSize.ToString()
        ItemCount        = $Stats.ItemCount
        ArchiveEnabled   = if ($_.ArchiveStatus -eq "Active") {"Yes"} else {"No"}
        ArchiveSize      = if ($ArchiveStats) { $ArchiveStats.TotalItemSize.ToString() } else {"N/A"}
    }
}

$MailboxData | Export-Csv -Path "S:\Tech Team\Benjamin\Scripts\MailboxReport.csv" -NoTypeInformation
