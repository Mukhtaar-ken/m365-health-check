function Test-RecoverableItemsQuota {
    <#
    .SYNOPSIS
        Flags mailboxes whose Recoverable Items folder is close to its quota.

    .DESCRIPTION
        TODO (you): explain in 2-3 lines what Recoverable Items is and why a
        full one is a problem (hint: what happens to deletes and holds?).

        This function only *evaluates* data. It never connects to Exchange,
        so it can be tested with fake objects.

    .PARAMETER Mailbox
        Objects with DisplayName, RecoverableItemsSizeBytes, RecoverableItemsQuotaBytes.

    .EXAMPLE
        Get-Content .\tests\fixtures\mailboxes.json | ConvertFrom-Json | Test-RecoverableItemsQuota
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
        # Runs once for each batch of mailboxes that comes down the pipeline.
        foreach ($mb in $Mailbox) {

            # No quota means we can't work out a percentage (and can't divide by zero).
            # No size means the data is missing, which must never be reported as 0% and OK.
            if (-not $mb.RecoverableItemsQuotaBytes -or $null -eq $mb.RecoverableItemsSizeBytes) {
                [pscustomobject]@{
                    Check       = 'RecoverableItemsQuota'
                    Target      = $mb.DisplayName
                    Status      = 'Unknown'
                    PercentUsed = $null
                    Detail      = 'The Recoverable Items size or quota is missing, so usage could not be checked.'
                }
                continue
            }

            $percent = [math]::Round($mb.RecoverableItemsSizeBytes / $mb.RecoverableItemsQuotaBytes * 100, 1)

            if ($percent -ge $CriticalPercent)    { $status = 'Critical' }
            elseif ($percent -ge $WarningPercent) { $status = 'Warning' }
            else                                  { $status = 'OK' }

            [pscustomobject]@{
                Check       = 'RecoverableItemsQuota'
                Target      = $mb.DisplayName
                Status      = $status
                PercentUsed = $percent
                Detail      = "Recoverable Items is $percent% full."
            }
        }
    }
}
