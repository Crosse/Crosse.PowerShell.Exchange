################################################################################
# 
# $URL$
# $Author$
# $Date$
# $Rev$
# 
# DESCRIPTION:  Provisions users in Exchange
# 
# 
# Copyright (c) 2009,2010 Seth Wright <wrightst@jmu.edu>
#
# Permission to use, copy, modify, and distribute this software for any
# purpose with or without fee is hereby granted, provided that the above
# copyright notice and this permission notice appear in all copies.
#
# THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
# WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
# MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
# ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
# WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
# ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
# OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
#
################################################################################

param ( $User="",
        $Server="localhost",
        [switch]$Automated=$false,
        [switch]$Mailbox=$true,
        [string]$ExternalEmailAddress=$null,
        [switch]$Force=$false,
        [switch]$Verbose=$false,
        $EmailAddresses=$null,
        [string]$RetryableErrorsFilePath=$null,
        [string]$DomainController=$null,
        [System.Collections.Hashtable]$Databases=$null,
        [switch]$Legacy=$false,
        $inputObject=$null )

# This section executes only once, before the pipeline.
BEGIN {
    if ($inputObject) {
        Write-Output $inputObject | &($MyInvocation.InvocationName)
        break
    }

    if ($Mailbox) {
        $srv = Get-ExchangeServer $Server -ErrorAction SilentlyContinue
        if ($srv -eq $null) {
            Write-Output "ERROR: Could not find Exchange Server $Server"
        }
    }

    if ([String]::IsNullOrEmpty($DomainController)) {
        $DomainController = (gc Env:\LOGONSERVER).Replace('\', '')
        if ([String]::IsNullOrEmpty($DomainController)) {
            Write-Output "ERROR:  Could not determine the local computer's logon server!"
            return
        }
    }

    if ($Verbose) {
        Write-Host "Using Domain Controller $DomainController"
    }

    $cwd = [System.IO.Path]::GetDirectoryName(($MyInvocation.MyCommand).Definition)

    $start = Get-Date
    if ([String]::IsNullOrEmpty($RetryableErrorsFilePath) -eq $false -or 
        [System.IO.File]::Exists($RetryableErrorsFilePath) -eq $false) {
            $RetryableErrorsFilePath = [System.IO.Path]::Combine($(Get-Location), "provisioning_errors_$($start.Ticks).csv")
    }

    if ($Mailbox -eq $true) {
        if ($Legacy) {
            if ($Databases -eq $null) {
                $Databases = & "$cwd\Get-BestDatabase.ps1" -Server $Server -Single:$false
                if ($Databases -eq $null) {
                    Write-Output "ERROR: Could not enumerate databases!"
                    return
                }
                if ($Verbose -eq $true) {
                    Write-Output "Found $($Databases.Count) databases."
                }
            }
        }
    } else {
        if ([String]::IsNullOrEmpty($ExternalEmailAddress)) {
            Write-Output "ERROR:  No ExternalEmailAddress given, and Mailbox is false"
            return
        }
    }

    $exitCode = 0
} # end 'BEGIN{}'

# This section executes for each object in the pipeline.
PROCESS {
    if ( !($_) -and !($User) ) { 
        Write-Output "ERROR: No user given."
        return
    }

    if ($_) { $User = $_ }

    # Was a username passed to us?  If not, bail.
    if ([String]::IsNullOrEmpty($User)) { 
        Write-Output "USAGE:  Provision-User -User `$User"
        return
    }

    $objUser = Get-User $User -ErrorAction SilentlyContinue

    if (!($objUser)) {
        Write-Output "$User`tis not a valid user in Active Directory."
        if ($Automated) {
            Out-File -NoClobber -Append -FilePath $RetryableErrorsFilePath -InputObject "$User,$start,No AD Account"
        }
        $exitCode += 1
        return
    } else { 
        switch ($objUser.RecipientTypeDetails) {
            'User' { break }
            'MailUser' {
                if (!$Mailbox) {
                    Write-Output "$($objUser.SamAccountName)`tis already a MailUser"
                    return
                    break
                }
            }
            'UserMailbox' {
                if ($Mailbox) {
                    Write-Output "$($objUser.SamAccountName)`talready has a mailbox"
                } else {
                    Write-Output "$($objUser.SamAccountName)`tis a Mailbox, refusing to enable as MailUser instead"
                    $exitCode += 1
                }
                return
            }
            'DisabledUser' {
                if ($Mailbox) {
                    Write-Output "$($objUser.SamAccountName)`tis disabled, refusing to create mailbox"
                    if ($Automated) {
                        Out-File -Append -FilePath $RetryableErrorsFilePath -InputObject "$User,$start,Disabled AD Account"
                    }
                    $exitCode += 1
                    return
                }
                break;
            }
            default {
                Write-Output "$($objUser.SamAccountName)`tis a $($objUser.RecipientTypeDetails) object, refusing to provision"
                $exitCode += 1
                return
            }
        }
    }

    # Save this off because Exchange blanks it out...
    $displayNamePrintable = $objUser.SimpleDisplayName

    if ($Mailbox) {
        # Don't auto-create mailboxes for users in the Students OU
        if ($objUser.DistinguishedName -match 'Student' -and $Force -eq $false) {
            Write-Output "$($User)`tis listed as a student, refusing to create mailbox"
            if ($Automated) {
                Out-File -Append -FilePath $RetryableErrorsFilePath -InputObject "$User,$start,Student"
            }
            # Don't mark this as being an error.
            #$exitCode += 1
            return
        }

        if ($Legacy) {
            $candidate = $null
            foreach ($db in $Databases.Keys) {
                if ($candidate -eq $null) {
                    $candidate = $db
                } else {
                    if ($Databases[$db] -lt $Databases[$candidate]) {
                        $candidate = $db
                    }
                }
            }

            if ($Verbose) {
                Write-Output "Assigning $($objUser.SamAccountName) to database $candidate"
            }
        }

        # If the user is a MailUser already, remove the Exchange bits first
        if ($objUser.RecipientTypeDetails -match 'MailUser') {
            & "$cwd\Deprovision-User.ps1" -User $objUser.DistinguishedName -Confirm:$false -DomainController $DomainController
            if ($LASTEXITCODE -gt 0) {
                Write-Output "An error occurred; refusing to create mailbox."
                if ($Automated) {
                    Out-File -Append -FilePath $RetryableErrorsFilePath -InputObject "$User,$start,Deprovisioning Error"
                }
                $exitCode += 1
                return
            }
        }

        # Enable the mailbox
        $Error.Clear()
        if ($Legacy) {
            Enable-Mailbox -Database "$($candidate)" -Identity $objUser.Identity `
                -ManagedFolderMailboxPolicy "Default Managed Folder Policy" `
                -ManagedFolderMailboxPolicyAllowed:$true `
                -DomainController $DomainController -ErrorAction SilentlyContinue

            # Increment the running mailbox total for the candidate database.
            $Databases[$candidate]++
        } else { 
            Enable-Mailbox -Identity $objUser.Identity `
                -ManagedFolderMailboxPolicy "Default Managed Folder Policy" `
                -ManagedFolderMailboxPolicyAllowed:$true `
                -DomainController $DomainController -ErrorAction SilentlyContinue
        }

        if ($Error[0] -ne $null) {
            Write-Output $Error[0]
            if ($Automated) {
                Out-File -Append -FilePath $RetryableErrorsFilePath -InputObject "$User,$start,$($Error[0].ToString())"
            }
            $exitCode += 1
            return
        } 

    } else {
        # The user should be enabled as a MailUser instead of a Mailbox.
        $Error.Clear()
        Enable-MailUser -Identity $objUser -ExternalEmailAddress $ExternalEmailAddress `
            -DomainController $DomainController

        if ($Error[0] -ne $null) {
            $exitCode += 1
            Write-Output $Error[0]
            if ($Automated) {
                Out-File -Append -FilePath $RetryableErrorsFilePath -InputObject "$User,$start,$($Error[0].ToString())"
            }
            return
        } 
    }

    # No error, so set the SimpleDisplayName now that Exchange has 
    # helpfully removed it.
    if ($Verbose) {
        Write-Output "Resetting $($objUser.SamAccountName)'s SimpleDisplayName to `"$displayNamePrintable`""
    }
    $error.Clear()
    Set-User $objUser -SimpleDisplayName "$($displayNamePrintable)" `
        -DomainController $DomainController -ErrorAction SilentlyContinue

    if (![String]::IsNullOrEmpty($error[0])) {
        Write-Output $error[0]
    }

    $EmailAddresses
    if ($EmailAddresses -eq $null) {
        if ($Verbose) {
            Write-Output "Not setting addresses"
        }
    } else {
        if ($Verbose) {
            Write-Output "Explicitly adding the following addresses to $($objUser.SamAccountName)'s EmailAddresses collection:"
            $EmailAddresses
        }
        $error.Clear()
        if ($Mailbox) {
            $addrs = (Get-Mailbox -Identity $objUser.Identity).EmailAddresses
            foreach ($addr in $addrs) { 
                if (!$EmailAddresses.Contains($addr)) {
                    $EmailAddresses.Add($addr)
                }
            }

            Set-Mailbox -Identity $objUser.Identity `
                -EmailAddressPolicyEnabled:$false `
                -EmailAddresses $EmailAddresses `
                -DomainController $DomainController
            Set-Mailbox -Identity $objUser.Identity `
                -EmailAddressPolicyEnabled:$true `
                -DomainController $DomainController
            if (![String]::IsNullOrEmpty($error[0])) {
                Write-Output $error[0]
            }
        } else {
            $addrs = (Get-MailUser -Identity $objUser.Identity).EmailAddresses
            foreach ($addr in $addrs) { 
                if (!$EmailAddresses.Contains($addr)) {
                    $EmailAddresses.Add($addr)
                }
            }

            Set-MailUser -Identity $objUser.Identity `
                -EmailAddressPolicyEnabled:$false `
                -EmailAddresses $EmailAddresses `
                -DomainController $DomainController
            Set-MailUser -Identity $objUser.Identity `
                -EmailAddressPolicyEnabled:$true `
                -DomainController $DomainController
            if (![String]::IsNullOrEmpty($error[0])) {
                Write-Output $error[0]
            }
        }
    }
} # end 'PROCESS{}'

# This section executes only once, after the pipeline.
END {
    exit $exitCode
} # end 'END{}'

