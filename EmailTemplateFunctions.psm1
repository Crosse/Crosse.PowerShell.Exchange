function Resolve-TemplatedEmail {
    param (
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [ValidateScript({ (Test-Path $_) })]
            [IO.FileInfo]
            $FilePath,

            [Parameter(Mandatory=$true)]
            [System.Collections.Hashtable]
            $TemplateSubstitutions,

            [switch]
            $ToBase64String
          )

    PROCESS {
        $Body = Get-Content $FilePath
        $Body = [String]::Join("`n", $Body)

        if ([String]::IsNullOrEmpty($Body)) {
            Write-Error "Template file $template either does not exist or is empty!"
            return
        }

        foreach ($key in $TemplateSubstitutions.Keys) {
            $Body = $Body.Replace($key, $TemplateSubstitutions[$key])
        }

        if ($ToBase64String) {
            $encoder = New-Object System.Text.ASCIIEncoding
            $Body = $encoder.GetBytes($Body)
            $Body = [Convert]::ToBase64String($Body)
        }
        return $Body
    }
}

function New-EmailDetailsObject {
    param (
            $Identity,
            $Address,
            $MessageBody,
            $Subject
          )

    return New-Object PSObject -Property @{
        Identity    = $Identity
        Address     = $Address
        MessageBody = $MessageBody
        Subject     = $Subject
    }
}

function Send-ProvisioningNotification {
    [CmdletBinding()]
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [ValidateScript({ (Test-Path $_) })]
            [IO.FileInfo]
            $FilePath,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [System.Net.Mail.MailAddress]
            $From,

            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $SmtpServer = $PSEmailServer,

            [switch]
            $UseSsl
          )
    PROCESS {
        $notifies = Import-Csv $FilePath

        foreach ($notify in $notifies) {
            if ([String]::IsNullOrEmpty($notify.MessageBody) -eq $false) {
                $subject = $notify.Subject
                $toaddr = $notify.Address
                $encoder = New-Object System.Text.ASCIIEncoding
                $body = $encoder.GetString([Convert]::FromBase64String($notify.MessageBody))

                Write-Verbose "Sending email to $toaddr"
                Send-MailMessage -To $toaddr -From $From -Body $body -BodyAsHtml `
                                 -Subject $subject -SmtpServer $SmtpServer -UseSsl:$UseSsl
            }
        }
    }
}
