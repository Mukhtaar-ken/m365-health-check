# Run with:  Invoke-Pester .\tests -Output Detailed   (needs Pester 5)
# Reports are written to TestDrive:, a temporary folder Pester creates and deletes for each test file.

BeforeAll {
    Import-Module "$PSScriptRoot\..\src\M365HealthCheck\M365HealthCheck.psd1" -Force
    $mailboxes = Get-Content "$PSScriptRoot\fixtures\mailboxes.json" -Raw | ConvertFrom-Json
}

Describe 'Invoke-M365HealthCheck' {

    Context 'with the fixture mailboxes' {
        BeforeAll {
            $outDir  = Join-Path $TestDrive 'report'
            $results = $mailboxes | Invoke-M365HealthCheck -OutputPath $outDir -InformationAction SilentlyContinue
        }

        It 'runs every check against every mailbox' {
            # 3 mailboxes x 2 checks
            $results.Count | Should -Be 6
            ($results.Check | Sort-Object -Unique) | Should -Be @('MailboxArchive', 'RecoverableItemsQuota')
        }

        It 'puts the worst results first' {
            ($results.Status -join ',') | Should -Be 'Critical,Warning,Warning,OK,OK,OK'
        }

        It 'creates the output folder and writes a CSV with every result' {
            $csv = Get-ChildItem -Path $outDir -Filter '*.csv'
            $csv | Should -HaveCount 1
            (Import-Csv -Path $csv.FullName) | Should -HaveCount 6
        }

        It 'writes an HTML report with a summary line' {
            $html = Get-ChildItem -Path $outDir -Filter '*.html'
            $html | Should -HaveCount 1
            Get-Content -Path $html.FullName -Raw | Should -Match 'Critical: 1 &middot; Warning: 2 &middot; OK: 3'
        }
    }

    It 'writes only the formats asked for' {
        $outDir = Join-Path $TestDrive 'csv-only'
        $null = $mailboxes | Invoke-M365HealthCheck -OutputPath $outDir -Format CSV -InformationAction SilentlyContinue

        Get-ChildItem -Path $outDir -Filter '*.csv'  | Should -HaveCount 1
        Get-ChildItem -Path $outDir -Filter '*.html' | Should -HaveCount 0
    }

    It 'collects data from Exchange when no mailboxes are piped in' {
        Mock Get-M365HealthData -ModuleName M365HealthCheck { $mailboxes[0] }

        $results = Invoke-M365HealthCheck -OutputPath (Join-Path $TestDrive 'collect') -InformationAction SilentlyContinue

        Should -Invoke Get-M365HealthData -ModuleName M365HealthCheck -Times 1 -Exactly
        $results.Count | Should -Be 2
    }
}
