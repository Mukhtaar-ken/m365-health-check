# Run with:  Invoke-Pester .\tests -Output Detailed   (needs Pester 5)

BeforeAll {
    Import-Module "$PSScriptRoot\..\src\M365HealthCheck\M365HealthCheck.psd1" -Force

    # Builds one fake mailbox. Sizes are in GB to keep the tests readable.
    function New-FakeMailbox ($Name, $UsedGB, $QuotaGB) {
        [pscustomobject]@{
            DisplayName                = $Name
            RecoverableItemsSizeBytes  = $UsedGB  * 1GB
            RecoverableItemsQuotaBytes = $QuotaGB * 1GB
        }
    }
}

Describe 'Test-RecoverableItemsQuota' {

    It 'returns OK when usage is under the warning threshold' {
        $result = New-FakeMailbox 'Alex Test' 10 30 | Test-RecoverableItemsQuota
        $result.Status | Should -Be 'OK'
    }

    It 'returns Warning between the warning and critical thresholds' {
        $result = New-FakeMailbox 'Sam Example' 26 30 | Test-RecoverableItemsQuota
        $result.Status | Should -Be 'Warning'
    }

    It 'returns Critical at or above the critical threshold' {
        $result = New-FakeMailbox 'Jo Sample' 30 30 | Test-RecoverableItemsQuota
        $result.Status | Should -Be 'Critical'
    }

    It 'handles several mailboxes from the pipeline' {
        $mailboxes = Get-Content "$PSScriptRoot\fixtures\mailboxes.json" -Raw | ConvertFrom-Json
        $results   = $mailboxes | Test-RecoverableItemsQuota
        $results.Count | Should -Be 3
    }

    It 'returns Unknown instead of throwing when the quota is 0' {
        $result = New-FakeMailbox 'No Quota' 5 0 | Test-RecoverableItemsQuota
        $result.Status | Should -Be 'Unknown'
    }

    It 'refuses a warning threshold that is not below the critical one' {
        { New-FakeMailbox 'Alex Test' 10 30 | Test-RecoverableItemsQuota -WarningPercent 90 -CriticalPercent 50 } |
            Should -Throw '*must be lower than*'
    }
}
