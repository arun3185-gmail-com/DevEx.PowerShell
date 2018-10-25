
[char[]] $Script:Alphas = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()

Function Convert-ExcelColumnNumberToName()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [int] $Number
    )

    if ($Number -lt 1)
    {
        throw New-Object System.ApplicationException("number must be greater than or equal to 1")
    }

    [int] $mod = $Number % 26
    [int] $coefOf26 = ($Number - $mod) / 26
    [int] $coefOf676 = ($Number - (26 * $coefOf26) - $mod) / 676
    [System.Text.StringBuilder] $colNameBuilder = New-Object System.Text.StringBuilder(3)

    if ($coefOf676 -eq 0) { $colNameBuilder.Append($Script:Alphas[25]) }
    elseif ($coefOf676 -gt 0) { $colNameBuilder.Append($Script:Alphas[$mod - 1]) }
    
    if ($coefOf26 -eq 0) { $colNameBuilder.Append($Script:Alphas[25]) }
    elseif ($coefOf26 -gt 0) { $colNameBuilder.Append($Script:Alphas[$mod - 1]) }

    if ($mod -eq 0) { $colNameBuilder.Append($Script:Alphas[25]) }
    elseif ($mod -gt 0) { $colNameBuilder.Append($Script:Alphas[$mod - 1]) }


    Write-Output $colNameBuilder.ToString()
}


Function Convert-ExcelColumnNameToNumber()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string] $Name
    )

    [int] $colNameLength = $Name.Length
    [int] $number = 0

    if ($colNameLength -ge 1) { $number +=       ([array]::IndexOf($Script:Alphas, $Name[$colNameLength - 1]) + 1) }
    if ($colNameLength -ge 2) { $number +=  26 * ([array]::IndexOf($Script:Alphas, $Name[$colNameLength - 2]) + 1) }
    if ($colNameLength -ge 3) { $number += 676 * ([array]::IndexOf($Script:Alphas, $Name[$colNameLength - 3]) + 1) }


    Write-Output $number
}