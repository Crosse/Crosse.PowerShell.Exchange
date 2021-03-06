function Get-RedirectRuleInformation {
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [object]
            $Identity
          )

    PROCESS {
        if ($Identity.GetType() -eq [Microsoft.Exchange.Data.Directory.Management.Mailbox]) {
            $mailbox = $Identity
        } else {
            $mailbox = Get-Mailbox $Identity
            if ($mailbox -eq $null) {
                Write-Error "$Identity does not have a mailbox"
                return
            }
        }

        $rules = Get-InboxRule -Mailbox $Identity
        $lang = (Get-MailboxRegionalConfiguration -Identity $Identity).Language

        if ($lang -eq $null) {
            $processed = $false
        } else {
            $processed = $true
        }

        $redirectTo = $null
        foreach ($rule in $rules) { 
            if ($rule.Enabled -eq $true -and $rule.RedirectTo -ne $null) { 
                foreach ($redirect in $rule.RedirectTo) {
                    $redirectTo = $redirect.Address
                }
            } 
        } 

        New-Object PSObject -Property @{ 
            "Identity"          = $mailbox.Name
            "MailboxAccessed"   = $processed
            "RedirectTo"        = $redirectTo }
    }
}
