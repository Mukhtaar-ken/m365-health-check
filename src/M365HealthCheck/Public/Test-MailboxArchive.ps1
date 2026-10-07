function Test-MailboxArchive {
    <#
    .SYNOPSIS
        Flags mailboxes with no archive that are filling up.

    .DESCRIPTION
        An online archive gives a mailbox somewhere to move older mail. Without one,
        a mailbox that reaches its send quota stops sending mail. No archive on its own
        is fine; no archive plus a nearly full mailbox is the problem.

        This function only *evaluates* data. It never connects to Exchange,
        so it can be tested with fake objects.

    .PARAMETER Mailbox
        Objects with DisplayName, ArchiveEnabled, MailboxSizeBytes, MailboxQuotaBytes.

    .EXAMPLE
        Get-Content .\tests\fixtures\mailboxes.json | ConvertFrom-Json | Test-MailboxArchive
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [object[]]$Mailbox,

        [ValidateRange(1, 100)]
        [int]$WarningPercent = 80,

        [ValidateRange(1, 100)]
        [int]$CriticalPercent = 95
    )

    begin {
        # Runs once, before any mailboxes. Stop straight away if the thresholds are the wrong way round.
        if ($WarningPercent -ge $CriticalPercent) {
            throw "WarningPercent ($WarningPercent) must be lower than CriticalPercent ($CriticalPercent)."
        }
    }

    process {
        foreach ($mb in $Mailbox) {

            # An archive means old mail has somewhere to go, so size doesn't matter here.
            if ($mb.ArchiveEnabled) {
                [pscustomobject]@{
                    Check       = 'MailboxArchive'
                    Target      = $mb.DisplayName
                    Status      = 'OK'
                    PercentUsed = $null
                    Detail      = 'Archive is enabled.'
                }
                continue
            }

            # No quota means we can't work out a percentage (and can't divide by zero).
            if (-not $mb.MailboxQuotaBytes) {
                [pscustomobject]@{
                    Check       = 'MailboxArchive'
                    Target      = $mb.DisplayName
                    Status      = 'Unknown'
                    PercentUsed = $null
                    Detail      = 'No archive, and no quota value, so usage could not be checked.'
                }
                continue
            }

            $percent = [math]::Round($mb.MailboxSizeBytes / $mb.MailboxQuotaBytes * 100, 1)

            if ($percent -ge $CriticalPercent)    { $status = 'Critical' }
            elseif ($percent -ge $WarningPercent) { $status = 'Warning' }
            else                                  { $status = 'OK' }

            [pscustomobject]@{
                Check       = 'MailboxArchive'
                Target      = $mb.DisplayName
                Status      = $status
                PercentUsed = $percent
                Detail      = "No archive, and the mailbox is $percent% full."
            }
        }
    }
}
