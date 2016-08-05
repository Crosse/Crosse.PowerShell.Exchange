function Connect-ToExchange {
    [CmdletBinding()]
    param (
            [System.Management.Automation.PSCredential]
            [ValidateNotNull()]
            $Credential = $(Get-Credential),

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $ConnectionUri,

            [switch]
            $AllowRedirection=$true,

            [ValidateNotNullOrEmpty()]
            [string]
            $Authentication,

            [switch]
            $ImportSession
          )

    Write-Verbose "Connecting to $ConnectionUri"
    $error.Clear()
    $session = New-PSSession    -ConfigurationName Microsoft.Exchange `
                                -ConnectionUri $ConnectionUri `
                                -Credential $Credential `
                                -Authentication $Authentication `
                                -AllowRedirection:$AllowRedirection `
                                -Verbose:$false `
                                -Name $Name

    if ([String]::IsNullOrEmpty($error[0]) -and $session -ne $null) {
        Write-Verbose "Connection successful."
        if ($ImportSession) {
            Write-Verbose "Importing session"
            Import-PSSession $session -AllowClobber
        }
    } else {
        Write-Error "Connection unsuccessful.  $($error[0])"
    }
    return $session
}

function Connect-ToLiveAtEdu {
    [CmdletBinding()]
    param (
            [ValidateNotNullOrEmpty()]
            [string]
            $ConnectionUri = "https://ps.outlook.com/PowerShell",

            [ValidateNotNull()]
            [System.Management.Automation.PSCredential]
            $Credential = $(Get-Credential),

            [switch]
            $AllowRedirection=$true,

            [switch]
            $ImportSession=$true
          )

    $session = Connect-ToExchange `
                    -ConnectionUri $ConnectionUri `
                    -Credential $Credential `
                    -AllowRedirection:$AllowRedirection `
                    -Authentication Basic

    if ($session -ne $null) {
        $session.Name = "Live@edu"

        if ((Test-Path Function:\Add-ShellType) -eq $true) {
            Add-ShellType "Live@edu"
        }

        if ($ImportSession) {
            Import-Session $session
        }
    }
}

function Disconnect-LiveAtEdu {
    Get-PSSession | Where-Object { $_.Name -match 'Live@edu' } | Remove-PSSession

    if ((Test-Path Function:\Remove-ShellType) -eq $true) {
        Remove-ShellType "Live@edu"
    }
}
