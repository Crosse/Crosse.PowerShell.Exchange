function Get-RedirectRuleInformation {
    param (
            [Parameter(Mandatory=$true,
                ValueFromPipeline=$true)]
            [ValidateNotNullOrEmpty()]
            [string]
            $Identity
          )

    PROCESS {
        $mailbox = Get-Mailbox $Identity 

        if ($mailbox -eq $null) {
            Write-Error "$Identity does not have a mailbox"
            return
        }

        $rules = Get-InboxRule -Mailbox $Identity
        $lang = (Get-MailboxRegionalConfiguration -Identity $Identity).Language

        if ($lang -eq $null) {
            $processed = $false
        } else {
            $processed = $true
        }

        if ($rules -eq $null) { 
            $redirectTo = $null
        } else {
            foreach ($rule in $rules) { 
                if ($rule.Enabled -eq $true -and $rule.RedirectTo -ne $null) { 
                    foreach ($redirect in $rule.RedirectTo) {
                        $redirectTo = $redirect.Replace("`"", "'")
                    }
                } 
            } 
        }

        New-Object PSObject -Property @{ 
            "Identity"          = $mailbox.Name
            "MailboxAccessed"   = $processed
            "RedirectTo"        = $redirectTo }
    }
}
