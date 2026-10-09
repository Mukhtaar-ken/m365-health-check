# Run with:  Invoke-Pester .\tests -Output Detailed   (needs Pester 5)

BeforeAll {
    Import-Module "$PSScriptRoot\..\src\M365HealthCheck\M365HealthCheck.psd1" -Force

    # Builds one fake mailbox. Sizes are in GB to keep the tests readable.
    function New-FakeMailbox ($Name, $Archive, $UsedGB, $QuotaGB) {
        [pscustomobject]@{
            DisplayName       = $Name
            ArchiveEnabled    = $Archive
            MailboxSizeBytes  = $UsedGB  * 1GB
            MailboxQuotaBytes = $QuotaGB * 1GB
        }
    }
}

Describe 'Test-MailboxArchive' {

    It 'returns OK when an archive is enabled, even if the mailbox is full' {
        $result = New-FakeMailbox 'Alex Test' $true 50 50 | Test-MailboxArchive
        $result.Status | Should -Be 'OK'
    }

    It 'returns OK with no archive when the mailbox is under the warning threshold' {
        $result = New-FakeMailbox 'Jo Sample' $false 10 50 | Test-MailboxArchive
        $result.Status | Should -Be 'OK'
    }

    It 'returns Warning with no archive between the warning and critical thresholds' {
        $result = New-FakeMailbox 'Sam Example' $false 42 50 | Test-MailboxArchive
        $result.Status | Should -Be 'Warning'
    }

    It 'returns Critical with no archive at or above the critical threshold' {
        $result = New-FakeMailbox 'Sam Example' $false 49 50 | Test-MailboxArchive
        $result.Status | Should -Be 'Critical'
    }

    It 'returns Unknown instead of throwing when there is no archive and the quota is 0' {
        $result = New-FakeMailbox 'No Quota' $false 5 0 | Test-MailboxArchive
        $result.Status | Should -Be 'Unknown'
    }

    It 'returns Unknown, not OK, when the mailbox size is missing' {
        $mb = New-FakeMailbox 'No Size' $false 0 50
        $mb.MailboxSizeBytes = $null
        ($mb | Test-MailboxArchive).Status | Should -Be 'Unknown'
    }

    It 'handles several mailboxes from the pipeline' {
        $mailboxes = Get-Content "$PSScriptRoot\fixtures\mailboxes.json" -Raw | ConvertFrom-Json
        $results   = $mailboxes | Test-MailboxArchive
        ($results.Status -join ',') | Should -Be 'OK,Warning,OK'
    }

    It 'refuses a warning threshold that is not below the critical one' {
        { New-FakeMailbox 'Alex Test' $false 10 50 | Test-MailboxArchive -WarningPercent 90 -CriticalPercent 50 } |
            Should -Throw '*must be lower than*'
    }
}
