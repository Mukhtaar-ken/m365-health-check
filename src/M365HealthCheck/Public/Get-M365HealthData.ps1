function Get-M365HealthData {
    <#
    .SYNOPSIS
        Collects mailbox data from Exchange Online in the shape the Test-* checks expect.

    .DESCRIPTION
        This is the only function in the module that talks to Exchange. It reads each
        user mailbox, its size, and its Recoverable Items folder, and turns Exchange's
        text sizes into plain byte counts. The checks then evaluate the result.

        It does not sign in for you. Run Connect-ExchangeOnline first, so how you
        authenticate (interactive, certificate, managed identity) stays your choice.

        Makes three calls per mailbox, so it is slow on large tenants.

    .EXAMPLE
        Connect-ExchangeOnline
        $data = Get-M365HealthData
        $data | Test-RecoverableItemsQuota
        $data | Test-MailboxArchive
    #>
    [CmdletBinding()]
    param()

    # Fail early with a clear message, rather than a confusing error halfway through.
    try   { $connection = Get-ConnectionInformation -ErrorAction Stop }
    catch { $connection = $null }

    if (-not $connection) {
        throw 'Not connected to Exchange Online. Install ExchangeOnlineManagement and run Connect-ExchangeOnline first.'
    }

    $mailboxes = Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited `
        -Properties ArchiveStatus, ProhibitSendQuota, RecoverableItemsQuota

    foreach ($mb in $mailboxes) {
        Write-Verbose "Reading $($mb.UserPrincipalName)"

        $stats = Get-EXOMailboxStatistics -Identity $mb.ExchangeGuid

        # The Recoverable Items scope returns several folders; the root one holds the total.
        $riRoot = Get-EXOMailboxFolderStatistics -Identity $mb.ExchangeGuid -FolderScope RecoverableItems |
            Where-Object FolderType -eq 'RecoverableItemsRoot'

        [pscustomobject]@{
            DisplayName                = $mb.DisplayName
            UserPrincipalName          = $mb.UserPrincipalName
            ArchiveEnabled             = $mb.ArchiveStatus -eq 'Active'
            MailboxSizeBytes           = ConvertTo-ByteCount $stats.TotalItemSize
            MailboxQuotaBytes          = ConvertTo-ByteCount $mb.ProhibitSendQuota
            RecoverableItemsSizeBytes  = ConvertTo-ByteCount $riRoot.FolderAndSubfolderSize
            RecoverableItemsQuotaBytes = ConvertTo-ByteCount $mb.RecoverableItemsQuota
        }
    }
}
