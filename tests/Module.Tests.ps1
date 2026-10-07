# Run with:  Invoke-Pester .\tests -Output Detailed   (needs Pester 5)
# Checks the module as a whole: the manifest is valid, and it exports exactly the Public functions.

BeforeAll {
    $manifestPath = "$PSScriptRoot\..\src\M365HealthCheck\M365HealthCheck.psd1"
    $publicDir    = "$PSScriptRoot\..\src\M365HealthCheck\Public"
}

Describe 'M365HealthCheck module' {

    It 'has a valid manifest' {
        { Test-ModuleManifest -Path $manifestPath -ErrorAction Stop } | Should -Not -Throw
    }

    It 'lists every Public function in the manifest' {
        # Catches the easy mistake: adding a new check file but forgetting FunctionsToExport.
        $manifest = Import-PowerShellDataFile -Path $manifestPath
        $files    = (Get-ChildItem -Path "$publicDir\*.ps1").BaseName

        @($manifest.FunctionsToExport) | Sort-Object | Should -Be ($files | Sort-Object)
    }

    It 'exports only the Public functions when imported' {
        $module = Import-Module $manifestPath -Force -PassThru
        $files  = (Get-ChildItem -Path "$publicDir\*.ps1").BaseName

        @($module.ExportedFunctions.Keys) | Sort-Object | Should -Be ($files | Sort-Object)
    }
}
