function New-ExchangeSession {
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Hostname,

            [Parameter(Mandatory=$true)]
            [System.Management.Automation.PSCredential]
            $Credential,

            [Parameter(Mandatory=$false)]
            $UseSsl=$false,

            [Parameter(Mandatory=$false)]
            $Authentication="Basic"
          )

    New-PSSession   -ConfigurationName Microsoft.Exchange `
                    -ConnectionUri "https://$($Hostname)/PowerShell" `
                    -Credential $Credential `
                    -Authentication $Authentication `
                    -AllowRedirection:$true
}
