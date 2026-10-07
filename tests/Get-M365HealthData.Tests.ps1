# Run with:  Invoke-Pester .\tests -Output Detailed   (needs Pester 5)
# Never connects to Exchange. The Exchange commands are replaced with fakes (mocks) that return made-up data.

BeforeAll {
    Import-Module "$PSScriptRoot\..\src\M365HealthCheck\M365HealthCheck.psd1" -Force

    # Pester can only mock a command that exists. The ExchangeOnlineManagement module
    # isn't installed on the CI machine, so create empty stand-ins for its commands.
    function global:Get-ConnectionInformation {}
    function global:Get-EXOMailbox {}
    function global:Get-EXOMailboxStatistics {}
    function global:Get-EXOMailboxFolderStatistics {}
}

AfterAll {
    Remove-Item -Path 'Function:\Get-ConnectionInformation', 'Function:\Get-EXOMailbox',
        'Function:\Get-EXOMailboxStatistics', 'Function:\Get-EXOMailboxFolderStatistics' -ErrorAction SilentlyContinue
}

Describe 'ConvertTo-ByteCount' {
    # A Private function, so it has to be tested from inside the module.

    It 'reads the byte count from an Exchange size' {
        InModuleScope M365HealthCheck { ConvertTo-ByteCount '1.2 GB (1,288,490,189 bytes)' } | Should -Be 1288490189
    }

    It 'returns nothing for Unlimited' {
        InModuleScope M365HealthCheck { ConvertTo-ByteCount 'Unlimited' } | Should -BeNullOrEmpty
    }

    It 'returns nothing for a blank value' {
        InModuleScope M365HealthCheck { ConvertTo-ByteCount $null } | Should -BeNullOrEmpty
    }
}

Describe 'Get-M365HealthData' {

    Context 'when not connected' {
        It 'stops with a clear message' {
            Mock Get-ConnectionInformation -ModuleName M365HealthCheck { $null }
            { Get-M365HealthData } | Should -Throw '*Connect-ExchangeOnline*'
        }
    }

    Context 'when connected' {
        BeforeAll {
            Mock Get-ConnectionInformation -ModuleName M365HealthCheck { [pscustomobject]@{ State = 'Connected' } }

            Mock Get-EXOMailbox -ModuleName M365HealthCheck {
                [pscustomobject]@{
                    DisplayName           = 'Sam Example'
                    UserPrincipalName     = 'sam@contoso.example'
                    ExchangeGuid          = [guid]'11111111-1111-1111-1111-111111111111'
                    ArchiveStatus         = 'None'
                    ProhibitSendQuota     = '49.5 GB (53,150,220,288 bytes)'
                    RecoverableItemsQuota = '30 GB (32,212,254,720 bytes)'
                }
            }

            Mock Get-EXOMailboxStatistics -ModuleName M365HealthCheck {
                [pscustomobject]@{ TotalItemSize = '45 GB (48,318,382,080 bytes)' }
            }

            # Two folders come back, like the real command; only the root should be used.
            # The subfolder is first on purpose, so "just take the first one" would fail the test.
            Mock Get-EXOMailboxFolderStatistics -ModuleName M365HealthCheck {
                [pscustomobject]@{ FolderType = 'Deletions';            FolderAndSubfolderSize = '1 GB (1,073,741,824 bytes)' }
                [pscustomobject]@{ FolderType = 'RecoverableItemsRoot'; FolderAndSubfolderSize = '26 GB (27,917,287,424 bytes)' }
            }

            $result = Get-M365HealthData
        }

        It 'turns Exchange sizes into byte counts' {
            $result.MailboxSizeBytes           | Should -Be 48318382080
            $result.MailboxQuotaBytes          | Should -Be 53150220288
            $result.RecoverableItemsQuotaBytes | Should -Be 32212254720
        }

        It 'uses the Recoverable Items root folder, not a subfolder' {
            $result.RecoverableItemsSizeBytes | Should -Be 27917287424
        }

        It 'treats ArchiveStatus None as no archive' {
            $result.ArchiveEnabled | Should -BeFalse
        }

        It 'produces objects both checks can read' {
            ($result | Test-MailboxArchive).Status        | Should -Be 'Warning'
            ($result | Test-RecoverableItemsQuota).Status | Should -Be 'Warning'
        }
    }
}
