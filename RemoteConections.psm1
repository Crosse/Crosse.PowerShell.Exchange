function ConnectTo-Exchange {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $ConnectionUri,

            [PSCredential]
            [ValidateNotNull()]
            $Credential = $(Get-Credential),

            [switch]
            $AllowRedirection=$true,

            [ValidateNotNullOrEmpty()]
            [string]
            $Authentication
          )

    Write-Verbose "Connecting to $ConnectionUri"
    $error.Clear()
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Credential $Credential -Authentication $Authentication -AllowRedirection:$AllowRedirection
    if ([String]::IsNullOrEmpty($error[0]) -or $session -ne $null) {
        Write-Verbose "Connection successful."
        Import-PSSession $session -AllowClobber
    } else {
        Write-Error "Connection unsuccessful.  $($error[0])"
    }
}

function ConnectTo-LiveAtEdu {
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $ConnectionUri = "https://ps.outlook.com/PowerShell",

            [PSCredential]
            [ValidateNotNull()]
            $Credential = $(Get-Credential),

            [switch]
            $AllowRedirection=$true
          )

    ConnectTo-Exchange @PSBoundParameters -Authentication Basic
}
