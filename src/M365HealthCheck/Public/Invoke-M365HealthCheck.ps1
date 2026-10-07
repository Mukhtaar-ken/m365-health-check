function Invoke-M365HealthCheck {
    <#
    .SYNOPSIS
        Runs every health check and saves the results as a CSV and/or HTML report.

    .DESCRIPTION
        Collects mailbox data (or uses the data you pipe in), runs every Test-* check
        in this module against it, sorts the results worst first, and writes the report.
        New checks are picked up automatically, as long as they are exported.

        Returns the results as well, so you can filter them further in PowerShell.

    .PARAMETER Mailbox
        Mailbox data to check. If left out, Get-M365HealthData collects it from
        Exchange Online, which needs Connect-ExchangeOnline first.

    .PARAMETER OutputPath
        Folder for the report files. Created if it doesn't exist.

    .PARAMETER Format
        CSV, HTML, or both (the default).

    .EXAMPLE
        Connect-ExchangeOnline
        Invoke-M365HealthCheck

    .EXAMPLE
        Get-Content .\tests\fixtures\mailboxes.json | ConvertFrom-Json | Invoke-M365HealthCheck -Format HTML
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        [object[]]$Mailbox,

        [string]$OutputPath = '.\output',

        [ValidateSet('CSV', 'HTML')]
        [string[]]$Format = @('CSV', 'HTML')
    )

    begin {
        # Show the "report saved" messages by default, unless the caller chose otherwise.
        if (-not $PSBoundParameters.ContainsKey('InformationAction')) {
            $InformationPreference = 'Continue'
        }

        $collected = [System.Collections.Generic.List[object]]::new()
    }

    process {
        # Piped mailboxes arrive in batches; gather them all before checking.
        foreach ($mb in $Mailbox) { $collected.Add($mb) }
    }

    end {
        if ($collected.Count -eq 0) {
            $collected.AddRange(@(Get-M365HealthData))
        }

        # Every exported Test-* function is a check. Adding a new one needs no change here.
        $checks = $MyInvocation.MyCommand.Module.ExportedFunctions.Keys | Where-Object { $_ -like 'Test-*' }

        # Worst first, so the problems are at the top of the report.
        $rank    = @{ Critical = 0; Warning = 1; Unknown = 2; OK = 3 }
        $results = foreach ($check in $checks) { $collected | & $check }
        $results = @($results | Sort-Object { $rank[$_.Status] }, Check, Target)

        if (-not (Test-Path -Path $OutputPath)) {
            New-Item -Path $OutputPath -ItemType Directory | Out-Null
        }
        $baseName = Join-Path $OutputPath ('health-check-{0:yyyyMMdd-HHmm}' -f (Get-Date))

        if ($Format -contains 'CSV') {
            $results | Export-Csv -Path "$baseName.csv" -NoTypeInformation
            Write-Information "CSV report: $baseName.csv"
        }

        if ($Format -contains 'HTML') {
            $summary = $results | Group-Object Status | Sort-Object { $rank[$_.Name] } |
                ForEach-Object { "$($_.Name): $($_.Count)" }

            $style = '<style>body{font-family:Segoe UI,sans-serif} table{border-collapse:collapse}
                td,th{border:1px solid #ccc;padding:4px 8px} th{background:#eee}</style>'

            $results |
                ConvertTo-Html -Title 'M365 health check' -Head $style `
                    -PreContent "<h1>M365 health check</h1><p>$(Get-Date -Format 'd MMM yyyy HH:mm') &middot; $($summary -join ' &middot; ')</p>" |
                Set-Content -Path "$baseName.html" -Encoding UTF8
            Write-Information "HTML report: $baseName.html"
        }

        $results
    }
}
