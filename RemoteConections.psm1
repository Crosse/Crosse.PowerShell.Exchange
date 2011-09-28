function ConnectTo-Exchange {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $ConnectionUri,

            [System.Management.Automation.PSCredential]
            [ValidateNotNull()]
            $Credential = $(Get-Credential),

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
                                -Verbose:$false

    if ([String]::IsNullOrEmpty($error[0]) -and $session -ne $null) {
        Write-Verbose "Connection successful."
        if ($ImportSession) {
            Write-Verbose "Importing session"
            Import-PSSession $session -AllowClobber
        }
    } else {
        Write-Error "Connection unsuccessful.  $($error[0])"
    }

    $session
}

function ConnectTo-LiveAtEdu {
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
            $ImportSession
          )

    ConnectTo-Exchange  -ConnectionUri $ConnectionUri `
                        -Credential $Credential `
                        -AllowRedirection:$AllowRedirection `
                        -Authentication Basic `
}
