
[string] $script:tabChar = [char]9
[string] $script:timeFormat = "yyyy-MM-dd HH:mm:ss.fff"
[string] $script:logFilePath = $null

function Initialize-DevExLogger
{
    [CmdletBinding()]
    Param
    (        
        [Parameter(Mandatory = $true)]
        [string] $FilePath,

        [Parameter(Mandatory = $false)]
        [bool] $Overwrite = $false
    )

    # Begin
    Begin
    {
        Write-Verbose -Message "Entering the BEGIN block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
    }

    # Process
    Process
    {
        Write-Verbose -Message "Entering the PROCESS block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."

        $script:logFilePath = $FilePath

        if ($Overwrite -and (Test-Path -Path $FilePath))
        {
            Remove-Item -Path $FilePath
        }

        Write-Verbose -Message "Leaving the PROCESS block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
    }

    # End
    End
    {
        Write-Verbose -Message "Entering the END block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
    }
}

function Write-DevExLog
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string] $Message
    )

    # Begin
    Begin
    {
        Write-Verbose -Message "Entering the BEGIN block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
    }

    # Process
    Process
    {
        Write-Verbose -Message "Entering the PROCESS block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."

        "[$(Get-Date -Format $script:timeFormat)]:$($script:tabChar)$($Message)" | Out-File -FilePath $script:logFilePath -Append

        Write-Verbose -Message "Leaving the PROCESS block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
    }
    
    # End
    End
    {
        Write-Verbose -Message "Entering the END block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
    }
}

function Out-DevExLog
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string] $Message
    )

    # Begin
    Begin
    {
        Write-Verbose -Message "Entering the BEGIN block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
    }

    # Process
    Process
    {
        Write-Verbose -Message "Entering the PROCESS block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."

        Write-DevExLog -Message $Message
        Write-Output $Message

        Write-Verbose -Message "Leaving the PROCESS block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
    }

    # End
    End
    {
        Write-Verbose -Message "Entering the END block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
    }
}
