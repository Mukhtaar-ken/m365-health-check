function ConvertTo-ByteCount {
    <#
    .SYNOPSIS
        Turns an Exchange size such as "1.2 GB (1,288,490,189 bytes)" into a number of bytes.

    .DESCRIPTION
        Exchange Online returns sizes and quotas as text. The exact byte count is the
        number in brackets, so we read that instead of trying to convert "1.2 GB".
        Returns $null for "Unlimited", blanks, or anything we can't read, which the
        checks report as Unknown.
    #>
    [CmdletBinding()]
    [OutputType([long])]
    param(
        [Parameter(ValueFromPipeline)]
        [AllowNull()]
        [object]$Value
    )

    process {
        # Some sizes (e.g. TotalItemSize) arrive wrapped as { IsUnlimited; Value }. Unwrap them first.
        if ($null -ne $Value -and $Value.PSObject.Properties['IsUnlimited']) {
            $Value = $Value.Value
        }

        # "$Value" turns either a string or an Exchange size object into the same text.
        if ("$Value" -match '\(([\d,]+) bytes\)') {
            [long]($Matches[1] -replace ',', '')
        }
        else {
            $null
        }
    }
}
